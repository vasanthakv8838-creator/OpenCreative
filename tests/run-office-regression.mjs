import http from "node:http";
import {createReadStream, promises as fs} from "node:fs";
import path from "node:path";
import {fileURLToPath} from "node:url";
import {chromium} from "playwright";

const root = path.resolve(fileURLToPath(new URL("../", import.meta.url)));
const mime = {
  ".html":"text/html; charset=utf-8",
  ".js":"text/javascript; charset=utf-8",
  ".css":"text/css; charset=utf-8",
  ".md":"text/markdown; charset=utf-8",
  ".json":"application/json; charset=utf-8",
  ".svg":"image/svg+xml"
};
const server = http.createServer(async (req, res) => {
  try {
    const pathname = decodeURIComponent(new URL(req.url || "/", "http://127.0.0.1").pathname);
    const relative = pathname === "/" ? "/index.html" : pathname;
    const target = path.resolve(root, "." + relative);
    if (target !== root && !target.startsWith(root + path.sep)) {
      res.writeHead(403).end("Forbidden");
      return;
    }
    const stat = await fs.stat(target);
    if (!stat.isFile()) throw new Error("Not a file");
    res.writeHead(200, {
      "Content-Type":mime[path.extname(target).toLowerCase()] || "application/octet-stream",
      "Content-Length":stat.size,
      "Cache-Control":"no-store",
      "X-Content-Type-Options":"nosniff"
    });
    createReadStream(target).pipe(res);
  } catch {
    res.writeHead(404, {"Content-Type":"text/plain; charset=utf-8"}).end("Not found");
  }
});

let browser;
try {
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  const url = "http://127.0.0.1:" + address.port + "/tests/office-interoperability.html";
  browser = await chromium.launch({headless:true});
  const context = await browser.newContext();
  const page = await context.newPage();
  const pageErrors = [];
  const cdnFailures = [];
  page.on("pageerror", error => pageErrors.push(error.message));
  page.on("requestfailed", request => {
    if (/cdn\.jsdelivr\.net/i.test(request.url())) cdnFailures.push(request.url() + " — " + (request.failure()?.errorText || "request failed"));
  });

  console.log("Opening " + url);
  await page.goto(url, {waitUntil:"domcontentloaded"});
  await page.waitForFunction(() => {
    const button = document.querySelector("#runTests");
    const status = document.querySelector("#status")?.textContent || "";
    return (button && !button.disabled) || !status.includes("Loading the productivity workspace");
  }, null, {timeout:120000});

  const runButton = page.locator("#runTests");
  if (!(await runButton.isEnabled())) {
    const status = await page.locator("#status").textContent();
    throw new Error("Office test engine did not become ready: " + status + (cdnFailures.length ? "\nCDN failures: " + cdnFailures.join("; ") : ""));
  }

  await runButton.click();
  await page.waitForFunction(() => {
    const status = document.querySelector("#status")?.textContent || "";
    return /^\d+ regression checks passed\.$/.test(status) || status.includes("Regression run failed");
  }, null, {timeout:120000});

  const status = (await page.locator("#status").textContent() || "").trim();
  const checks = await page.locator("#results li.pass").allTextContents();
  const failures = await page.locator("#results li.fail").allTextContents();
  if (failures.length || !/^\d+ regression checks passed\.$/.test(status)) {
    throw new Error(status + (failures.length ? "\n" + failures.join("\n") : ""));
  }
  if (pageErrors.length) throw new Error("Uncaught browser errors: " + pageErrors.join("; "));
  console.log(status);
  checks.forEach(check => console.log(check));
  console.log("Chromium Office regression run passed.");
} catch (error) {
  console.error("Chromium Office regression run failed:", error);
  process.exitCode = 1;
} finally {
  if (browser) await browser.close();
  if (server.listening) await new Promise(resolve => server.close(resolve));
}
