#!/usr/bin/env node

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const { google } = require("googleapis");
const { OAuth2Client } = require("google-auth-library");

function parseArgs(argv) {
  const out = {
    appDir: "",
    deploymentId: "",
    skipSync: false,
    mode: "versioned",
    settings: [],
    adminCredential: "test",
    storageState: "/home/mifasol/interior-app/tools/.auth/google_quote_admin.json"
  };
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === "--app-dir") out.appDir = argv[i + 1] || "";
    if (token === "--deployment-id") out.deploymentId = argv[i + 1] || "";
    if (token === "--skip-sync") out.skipSync = true;
    if (token === "--mode") out.mode = String(argv[i + 1] || "versioned").trim().toLowerCase();
    if (token === "--setting") out.settings.push(String(argv[i + 1] || ""));
    if (token === "--admin-credential") out.adminCredential = String(argv[i + 1] || "").trim();
    if (token === "--storage-state") out.storageState = String(argv[i + 1] || "").trim();
  }
  if (!out.appDir) {
    throw new Error("Usage: node tools/release_webapp.js --app-dir <quote_app|prequote_app> [--deployment-id <id>] [--mode <versioned|head>] [--setting key=value] [--admin-credential <value>] [--storage-state <path>] [--skip-sync]");
  }
  return out;
}

function loadClaspAuth() {
  const rcPath = path.join(os.homedir(), ".clasprc.json");
  const rc = JSON.parse(fs.readFileSync(rcPath, "utf8"));
  const client = new OAuth2Client(
    rc.oauth2ClientSettings.clientId,
    rc.oauth2ClientSettings.clientSecret,
    rc.oauth2ClientSettings.redirectUri
  );
  client.setCredentials(rc.token);
  return client;
}

function readScriptId(appDir) {
  const claspPath = path.join(appDir, ".clasp.json");
  const config = JSON.parse(fs.readFileSync(claspPath, "utf8"));
  if (!config.scriptId) throw new Error(".clasp.json missing scriptId");
  return String(config.scriptId).trim();
}

function runCommand(command, args, cwd) {
  const result = spawnSync(command, args, {
    cwd,
    stdio: "inherit",
    shell: false
  });
  if (result.status !== 0) {
    throw new Error([command].concat(args).join(" ") + " failed with exit code " + result.status);
  }
}

function pickDeployment(deployments, preferredId, mode) {
  const list = Array.isArray(deployments) ? deployments : [];
  if (preferredId) {
    const hit = list.find((item) => String(item.deploymentId || "") === preferredId);
    if (!hit) throw new Error("Deployment not found: " + preferredId);
    return hit;
  }

  const webApps = list.filter((item) =>
    Array.isArray(item.entryPoints) &&
    item.entryPoints.some((entry) => entry.entryPointType === "WEB_APP" && entry.webApp && entry.webApp.url)
  );
  const versioned = webApps
    .filter((item) => item.deploymentConfig && Number(item.deploymentConfig.versionNumber || 0) > 0)
    .sort((a, b) => Number(b.deploymentConfig.versionNumber || 0) - Number(a.deploymentConfig.versionNumber || 0));
  const heads = webApps.filter((item) => !item.deploymentConfig || !Number(item.deploymentConfig.versionNumber || 0));

  if (mode === "head") return heads[0] || webApps[0] || null;
  return versioned[0] || heads[0] || webApps[0] || null;
}

function getWebAppUrl(deployment) {
  const entry = (deployment.entryPoints || []).find((item) => item.entryPointType === "WEB_APP" && item.webApp && item.webApp.url);
  return entry ? String(entry.webApp.url || "").trim() : "";
}

async function ensureSettingsValues(scriptApi, scriptId, pairs, adminCredential) {
  const normalizedPairs = (pairs || [])
    .map((item) => ({
      key: String(item && item.key || "").trim(),
      value: String(item && item.value || "").trim()
    }))
    .filter((item) => item.key);

  const res = await scriptApi.scripts.run({
    scriptId,
    requestBody: {
      function: "applyReleaseSettings_ADMIN",
      parameters: [normalizedPairs, String(adminCredential || "").trim()],
      devMode: true
    }
  });

  if (res.data.error) {
    throw new Error(JSON.stringify(res.data.error));
  }

  return Object.assign({ keys: normalizedPairs.map((item) => item.key) }, res.data.response && res.data.response.result ? res.data.response.result : {});
}

function ensureSettingsValuesViaBrowser(repoRoot, appName, spreadsheetId, pairs, adminCredential, storageState) {
  const helper = path.join(repoRoot, "tools", "apps_script_browser_admin.py");
  const paramsSetup = JSON.stringify([spreadsheetId]);
  const paramsApply = JSON.stringify([pairs, String(adminCredential || "").trim()]);

  runCommand("python3", [
    helper,
    "--app", appName,
    "--function", "setupSpreadsheetIdFromManual",
    "--params-json", paramsSetup,
    "--storage-state", storageState
  ], repoRoot);

  const result = spawnSync("python3", [
    helper,
    "--app", appName,
    "--function", "applyReleaseSettings_ADMIN",
    "--params-json", paramsApply,
    "--storage-state", storageState
  ], {
    cwd: repoRoot,
    encoding: "utf8",
    shell: false
  });

  if (result.status !== 0) {
    throw new Error((result.stderr || result.stdout || "").trim() || "browser admin sync failed");
  }

  const payload = JSON.parse(String(result.stdout || "{}"));
  if (!payload.ok) {
    throw new Error(String(payload.error || "browser admin sync failed"));
  }
  return Object.assign({ fallback: "browser" }, payload.result || {});
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const repoRoot = path.resolve(__dirname, "..");
  const appDir = path.resolve(repoRoot, args.appDir);
  const appName = path.basename(appDir).replace(/_app$/, "");
  const scriptId = readScriptId(appDir);
  const auth = loadClaspAuth();
  const script = google.script({ version: "v1", auth });

  const beforeDeployments = await script.projects.deployments.list({ scriptId });
  const selected = pickDeployment(beforeDeployments.data.deployments, args.deploymentId, args.mode);
  if (!selected) throw new Error("No web app deployment found for " + args.appDir);

  runCommand("npx", ["clasp", "push", "-f"], appDir);

  if (args.mode !== "head") {
    runCommand("npx", [
      "clasp",
      "deploy",
      "--deploymentId",
      String(selected.deploymentId),
      "--description",
      "codex release " + new Date().toISOString()
    ], appDir);
  }

  const afterDeployments = await script.projects.deployments.list({ scriptId });
  const deployed = pickDeployment(afterDeployments.data.deployments, String(selected.deploymentId), args.mode);
  const project = await script.projects.get({ scriptId });
  const spreadsheetId = String(project.data.parentId || "").trim();
  const baseUrl = getWebAppUrl(deployed);
  let sync = { added: false, updated: false, skipped: true, warning: "" };

  if (!args.skipSync) {
    if (!spreadsheetId) throw new Error("Missing parent spreadsheet id for " + args.appDir);
    if (!baseUrl) throw new Error("Missing web app url for deployment " + selected.deploymentId);
    const extraSettings = args.settings
      .map((item) => {
        const idx = item.indexOf("=");
        if (idx < 0) return null;
        return { key: item.slice(0, idx), value: item.slice(idx + 1) };
      })
      .filter(Boolean);
    try {
      sync = await ensureSettingsValues(script, scriptId, [{ key: "base_url", value: baseUrl }].concat(extraSettings), args.adminCredential);
      sync.skipped = false;
    } catch (error) {
      try {
        sync = ensureSettingsValuesViaBrowser(
          repoRoot,
          appName,
          spreadsheetId,
          [{ key: "base_url", value: baseUrl }].concat(extraSettings),
          args.adminCredential,
          args.storageState
        );
        sync.skipped = false;
        sync.warning = String(error && error.message ? error.message : error);
      } catch (fallbackError) {
        sync.skipped = true;
        sync.warning = String(error && error.message ? error.message : error);
        sync.fallback_warning = String(fallbackError && fallbackError.message ? fallbackError.message : fallbackError);
      }
    }
  }

  const summary = {
    app_dir: args.appDir,
    mode: args.mode,
    script_id: scriptId,
    spreadsheet_id: spreadsheetId,
    deployment_id: String(selected.deploymentId),
    version_number: Number((deployed.deploymentConfig && deployed.deploymentConfig.versionNumber) || 0) || null,
    base_url: baseUrl,
    sync
  };

  process.stdout.write(JSON.stringify(summary, null, 2) + "\n");
}

main().catch((error) => {
  process.stderr.write(String(error && error.stack ? error.stack : error) + "\n");
  process.exit(1);
});
