'use strict';

// Invoke one native command without a shell. No dependency on a persistent PS session.
const fs = require('node:fs');
const path = require('node:path');
const { spawn, execFileSync } = require('node:child_process');
const root = path.resolve(__dirname, '../../..');

function withinRoot(value) {
  const absolute = path.resolve(root, value);
  const relative = path.relative(root, absolute);
  if (relative === '..' || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new Error('Path must stay inside the project workspace.');
  }
  return absolute;
}

function plan() {
  const config = JSON.parse(fs.readFileSync(path.join(__dirname, '../targets.json'), 'utf8'));
  const packages = {};
  for (const group of [config.framework, config.tooling, config.components]) {
    for (const name of group.packages) packages[name] = group.version;
  }
  Object.assign(packages, config.otherPackages);
  return { packages, migrations: config.migrationPackages.map(name => ({ name, to: packages[name] })) };
}

function npmEntry() {
  const candidates = [process.env.npm_execpath,
    path.join(path.dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js')];
  const isNpmEntry = file => file && file.endsWith('npm-cli.js') && fs.existsSync(file);
  const direct = candidates.find(isNpmEntry);
  if (direct) return direct;
  if (process.platform === 'win32') {
    const result = execFileSync('where.exe', ['npm.cmd'], { encoding: 'utf8', windowsHide: true });
    for (const file of result.trim().split(/\r?\n/)) {
      candidates.push(path.join(path.dirname(file), 'node_modules/npm/bin/npm-cli.js'));
    }
  }
  const entry = candidates.find(isNpmEntry);
  if (!entry) throw new Error('Cannot locate npm-cli.js. Inspect the approved Node/npm installation.');
  return entry;
}

async function main() {
  const [runArg, label, tool, ...args] = process.argv.slice(2);
  if (runArg === '--targets' && !label) {
    console.log(JSON.stringify(plan(), null, 2));
    return;
  }
  if (!runArg || !/^[a-z0-9][a-z0-9-]*$/.test(label || '')) {
    throw new Error('Usage: node run.cjs RUN_DIRECTORY LABEL npm|ng|node|git [arguments]; or --targets');
  }
  const cwd = fs.realpathSync(process.cwd());
  if (cwd.toLowerCase() !== fs.realpathSync(root).toLowerCase()) {
    throw new Error('Run from the project root.');
  }
  const runDirectory = withinRoot(runArg);
  const tmpDirectory = path.join(root, '.tmp');
  const relativeTmp = path.relative(tmpDirectory, runDirectory);
  if (!relativeTmp || relativeTmp === '..' || relativeTmp.startsWith(`..${path.sep}`)
      || path.isAbsolute(relativeTmp)) throw new Error('Logs must be in a subdirectory of .tmp.');
  const env = { ...process.env, CI: 'true', NG_CLI_ANALYTICS: 'false', NO_COLOR: '1' };
  let executable;
  let commandArgs;
  switch (tool) {
    case 'npm':
      executable = process.execPath;
      commandArgs = [npmEntry(), ...args];
      break;
    case 'ng': {
      const entry = path.join(root, 'node_modules/@angular/cli/bin/ng.js');
      if (!fs.existsSync(entry)) throw new Error('Local Angular CLI missing. Install the reviewed lock first.');
      // Only bypass CLI delegation/automatic temporary downloads, never peers or engines.
      env.NG_DISABLE_VERSION_CHECK = 'true';
      executable = process.execPath;
      commandArgs = [entry, ...args];
      break;
    }
    case 'node':
      executable = process.execPath;
      commandArgs = args;
      break;
    case 'git':
      executable = 'git';
      commandArgs = args;
      break;
    default: throw new Error('Unsupported tool. Use npm, ng, node or git.');
  }
  fs.mkdirSync(runDirectory, { recursive: true });
  const token = new Date().toISOString().replace(/[:.]/g, '-') + '-' + process.pid;
  const logPath = path.join(runDirectory, `${label}-${token}.log`);
  const log = fs.createWriteStream(logPath, { flags: 'wx' });
  const record = { label, tool, args, cwd, node: process.version,
    startedAt: new Date().toISOString(), log: path.relative(root, logPath) };
  const ledger = path.join(runDirectory, 'commands.jsonl');
  fs.appendFileSync(ledger, JSON.stringify({ ...record, status: 'started' }) + '\n');
  console.log(`${label}: ${tool} ${JSON.stringify(args)}\nLog: ${record.log}`);
  const child = spawn(executable, commandArgs, {
    cwd: root, env, windowsHide: true, shell: false, stdio: ['ignore', 'pipe', 'pipe']
  });
  let launchError;
  for (const stream of [child.stdout, child.stderr]) {
    stream.on('data', data => {
      log.write(data);
      // Relay both streams to stdout to avoid PS 5.1 turning warnings into native errors.
      process.stdout.write(data);
    });
  }
  child.on('error', error => { launchError = error.message; });
  await new Promise(resolve => child.on('close', (code, signal) => {
    const exitCode = Number.isInteger(code) ? code : 1;
    log.end(() => {
      const result = { ...record, status: exitCode === 0 ? 'passed' : 'failed',
        exitCode, signal, launchError, finishedAt: new Date().toISOString() };
      fs.appendFileSync(ledger, JSON.stringify(result) + '\n');
      console.log(`\n${label}: exit ${exitCode}${signal ? ` (${signal})` : ''}`);
      if (launchError) console.log(launchError);
      process.exitCode = exitCode;
      resolve();
    });
  }));
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
