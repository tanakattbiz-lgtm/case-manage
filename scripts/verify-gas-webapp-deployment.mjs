import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import { GoogleAuth } from 'google-auth-library';
import { google } from 'googleapis';

const DEFAULT_CLASP_OAUTH_CLIENT_ID =
  '1072944905499-vm2v2i5dvn0a0d2o4ca36i1vge8cvbn0.apps.googleusercontent.com';
const DEFAULT_CLASP_OAUTH_CLIENT_SECRET = 'eASZ9kFGC0g0JHKQ35WBQWft';

function fail(message) {
  console.error(message);
  process.exit(1);
}

function loadJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function normalizeDeploymentId(value) {
  if (!value) {
    return null;
  }

  const webAppUrlMatch = value.match(/\/s\/([^/]+)\/exec(?:[?#].*)?$/);
  return webAppUrlMatch ? webAppUrlMatch[1] : value;
}

function readClaspProjectConfig() {
  const projectConfigPath = path.resolve(process.cwd(), '.clasp.json');
  if (!fs.existsSync(projectConfigPath)) {
    return null;
  }

  return loadJson(projectConfigPath);
}

function loadStoredCredentials(filePath) {
  if (!fs.existsSync(filePath)) {
    fail(`clasp auth file was not found: ${filePath}`);
  }

  const store = loadJson(filePath);

  if (store.tokens) {
    if (store.tokens.default) {
      return store.tokens.default;
    }

    const firstStoredCredential = Object.values(store.tokens).find(Boolean);
    if (firstStoredCredential) {
      return firstStoredCredential;
    }
  }

  if (store.token && store.oauth2ClientSettings) {
    return {
      type: 'authorized_user',
      ...store.token,
      client_id: store.oauth2ClientSettings.clientId,
      client_secret: store.oauth2ClientSettings.clientSecret,
    };
  }

  if (store.access_token || store.refresh_token) {
    return {
      type: 'authorized_user',
      access_token: store.access_token,
      refresh_token: store.refresh_token,
      expiry_date: store.exprity_date,
      token_type: store.token_type,
      client_id: DEFAULT_CLASP_OAUTH_CLIENT_ID,
      client_secret: DEFAULT_CLASP_OAUTH_CLIENT_SECRET,
    };
  }

  fail('No usable clasp credentials were found in .clasprc.json.');
}

function readManifestWebAppConfig() {
  const manifestPath = path.resolve(process.cwd(), 'src', 'appsscript.json');
  if (!fs.existsSync(manifestPath)) {
    fail(`Apps Script manifest was not found: ${manifestPath}`);
  }

  try {
    const manifest = loadJson(manifestPath);
    if (!manifest.webapp) {
      fail('src/appsscript.json must define a "webapp" block.');
    }

    return manifest.webapp;
  } catch (error) {
    fail(`Failed to read src/appsscript.json: ${error.message}`);
  }
}

async function main() {
  const projectConfig = readClaspProjectConfig();
  const manifestWebAppConfig = readManifestWebAppConfig();
  const scriptId = process.env.GAS_SCRIPT_ID || projectConfig?.scriptId;
  const deploymentId = normalizeDeploymentId(process.env.GAS_WEBAPP_DEPLOYMENT_ID);
  const authFilePath =
    process.env.CLASP_AUTH_FILE || path.join(os.homedir(), '.clasprc.json');

  if (!scriptId) {
    fail('GAS_SCRIPT_ID is not set and .clasp.json has no scriptId.');
  }

  if (!deploymentId) {
    fail('GAS_WEBAPP_DEPLOYMENT_ID is not set.');
  }

  const storedCredentials = loadStoredCredentials(authFilePath);
  const auth = new GoogleAuth().fromJSON(storedCredentials);
  auth.setCredentials(storedCredentials);

  const script = google.script({ version: 'v1', auth });
  const response = await script.projects.deployments.get({
    scriptId,
    deploymentId,
  });

  const deployment = response.data || {};
  const entryPoints = deployment.entryPoints || [];
  const webAppEntry = entryPoints.find((entryPoint) => entryPoint.webApp);

  if (!webAppEntry?.webApp) {
    const entryPointTypes = entryPoints
      .map((entryPoint) => entryPoint.entryPointType)
      .filter(Boolean);
    const detail =
      entryPointTypes.length > 0
        ? ` Entry points: ${entryPointTypes.join(', ')}`
        : ' Entry points: none';
    fail(
      `Deployment ${deploymentId} is not a Web App deployment for script ${scriptId}.${detail}`,
    );
  }

  const config = webAppEntry.webApp.entryPointConfig || {};

  if (config.access !== manifestWebAppConfig.access) {
    fail(
      `Deployment ${deploymentId} access is ${config.access || 'undefined'}, expected ${manifestWebAppConfig.access}.`,
    );
  }

  if (config.executeAs !== manifestWebAppConfig.executeAs) {
    fail(
      `Deployment ${deploymentId} executeAs is ${config.executeAs || 'undefined'}, expected ${manifestWebAppConfig.executeAs}.`,
    );
  }

  console.log(`Verified Web App deployment: ${deploymentId}`);
  if (webAppEntry.webApp.url) {
    console.log(`Web App URL: ${webAppEntry.webApp.url}`);
  }
  if (config.access) {
    console.log(`Access: ${config.access}`);
  }
  if (config.executeAs) {
    console.log(`ExecuteAs: ${config.executeAs}`);
  }
}

main().catch((error) => {
  fail(error instanceof Error ? error.stack || error.message : String(error));
});
