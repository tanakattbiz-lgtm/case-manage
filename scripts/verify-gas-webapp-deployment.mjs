import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

function requireEnv(name) {
  const value = process.env[name];
  if (!value || !value.trim()) {
    throw new Error(`${name} is not set.`);
  }
  return value.trim();
}

async function runClasp(args) {
  const command = "node";
  const claspEntry = "./node_modules/@google/clasp/build/src/index.js";
  const { stdout, stderr } = await execFileAsync(command, [claspEntry, ...args], {
    env: process.env,
    maxBuffer: 10 * 1024 * 1024
  });

  if (stdout?.trim()) {
    console.log(stdout.trim());
  }
  if (stderr?.trim()) {
    console.error(stderr.trim());
  }

  return stdout ?? "";
}

async function main() {
  requireEnv("GAS_SCRIPT_ID");
  const deploymentId = requireEnv("GAS_WEBAPP_DEPLOYMENT_ID");

  console.log(`==> Checking deployment: ${deploymentId}`);
  const output = await runClasp(["list-deployments"]);

  if (!output.includes(deploymentId)) {
    throw new Error(
      [
        "GAS_WEBAPP_DEPLOYMENT_ID was not found in clasp list-deployments output.",
        `Expected deploymentId: ${deploymentId}`,
        "Check that the secret points to an existing Apps Script Web App deployment."
      ].join("\n")
    );
  }

  console.log("==> Deployment ID exists in this Apps Script project");
}

main().catch((error) => {
  console.error("Verification failed.");
  console.error(error instanceof Error ? error.stack ?? error.message : String(error));
  process.exit(1);
});
