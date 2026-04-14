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

async function runClasp(args, { allowFailure = false } = {}) {
  const command = "node";
  const claspEntry = "./node_modules/@google/clasp/build/src/index.js";

  try {
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

    return { stdout: stdout ?? "", stderr: stderr ?? "" };
  } catch (error) {
    const stdout = error.stdout ?? "";
    const stderr = error.stderr ?? "";
    if (stdout.trim()) {
      console.log(stdout.trim());
    }
    if (stderr.trim()) {
      console.error(stderr.trim());
    }
    if (allowFailure) {
      return { stdout, stderr, error };
    }
    throw error;
  }
}

function parseCreatedVersion(output) {
  const match = output.match(/Created version (\d+)/i);
  if (!match) {
    throw new Error(`Failed to parse version number from output: ${output}`);
  }
  return match[1];
}

async function main() {
  requireEnv("GAS_SCRIPT_ID");
  const deploymentId = requireEnv("GAS_WEBAPP_DEPLOYMENT_ID");
  const revision = process.env.GITHUB_SHA?.slice(0, 7) ?? "local";
  const description = `GitHub Actions ${revision}`;

  console.log("==> Pushing source to Apps Script");
  await runClasp(["push", "--force"]);

  console.log("==> Creating immutable version");
  const versionResult = await runClasp(["version", description]);
  const versionNumber = parseCreatedVersion(versionResult.stdout);

  console.log(`==> Redeploying existing Web App deployment: ${deploymentId}`);
  await runClasp(["redeploy", deploymentId, versionNumber, description]);

  console.log("==> Deployment completed");
  console.log(`deploymentId=${deploymentId}`);
  console.log(`version=${versionNumber}`);
}

main().catch((error) => {
  console.error("Deployment failed.");
  console.error(error instanceof Error ? error.stack ?? error.message : String(error));
  process.exit(1);
});
