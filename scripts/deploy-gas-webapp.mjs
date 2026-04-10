import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';

function fail(message) {
  console.error(message);
  process.exit(1);
}

function normalizeDeploymentId(value) {
  if (!value) {
    return null;
  }

  const webAppUrlMatch = value.match(/\/s\/([^/]+)\/exec(?:[?#].*)?$/);
  return webAppUrlMatch ? webAppUrlMatch[1] : value;
}

function runNodeScript(scriptPath, args = []) {
  const result = spawnSync(process.execPath, [scriptPath, ...args], {
    cwd: process.cwd(),
    env: process.env,
    stdio: 'inherit',
  });

  if (result.status !== 0) {
    fail(`Command failed: node ${path.relative(process.cwd(), scriptPath)}`);
  }
}

function runClasp(args) {
  const claspCliPath = path.resolve(
    process.cwd(),
    'node_modules',
    '@google',
    'clasp',
    'build',
    'src',
    'index.js',
  );

  if (!fs.existsSync(claspCliPath)) {
    fail(`clasp CLI was not found: ${claspCliPath}`);
  }

  const result = spawnSync(process.execPath, [claspCliPath, ...args], {
    cwd: process.cwd(),
    env: process.env,
    stdio: 'inherit',
  });

  if (result.status !== 0) {
    fail(`clasp command failed: clasp ${args.join(' ')}`);
  }
}

function ensureProjectConfig() {
  const scriptId = process.env.GAS_SCRIPT_ID;
  const projectConfigPath = path.resolve(process.cwd(), '.clasp.json');

  if (!scriptId && !fs.existsSync(projectConfigPath)) {
    fail('GAS_SCRIPT_ID is required when .clasp.json does not exist.');
  }

  if (!scriptId) {
    return;
  }

  const projectConfig = {
    scriptId,
    rootDir: 'src',
  };

  if (fs.existsSync(projectConfigPath)) {
    const existingConfig = JSON.parse(fs.readFileSync(projectConfigPath, 'utf8'));

    if (existingConfig.scriptId && existingConfig.scriptId !== scriptId) {
      fail(
        `.clasp.json scriptId (${existingConfig.scriptId}) does not match GAS_SCRIPT_ID (${scriptId}).`,
      );
    }
  }

  fs.writeFileSync(projectConfigPath, `${JSON.stringify(projectConfig, null, 2)}\n`, 'utf8');
}

function main() {
  const deploymentId = normalizeDeploymentId(process.env.GAS_WEBAPP_DEPLOYMENT_ID);
  const verifyScriptPath = path.resolve(
    process.cwd(),
    'scripts',
    'verify-gas-webapp-deployment.mjs',
  );

  if (!deploymentId) {
    fail('GAS_WEBAPP_DEPLOYMENT_ID is not set.');
  }

  ensureProjectConfig();

  runNodeScript(verifyScriptPath);
  runClasp(['show-file-status']);
  runClasp(['push', '--force']);
  runClasp([
    'update-deployment',
    deploymentId,
    '--description',
    `Web App deploy ${new Date().toISOString()}`,
  ]);
  runNodeScript(verifyScriptPath);
}

main();
