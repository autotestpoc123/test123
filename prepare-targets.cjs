'use strict';

// Prepare all mutually constrained direct dependencies in one manifest transaction.
// Default is a dry run. This script never installs packages or runs migrations.
const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { execFileSync } = require('node:child_process');
const root = path.resolve(__dirname, '../../..');
const runner = path.join(__dirname, 'run.cjs');
const plan = JSON.parse(execFileSync(process.execPath, [runner, '--targets'], { encoding: 'utf8' }));

function main() {
  const args = process.argv.slice(2);
  let write = false;
  let directory = '.';
  for (let index = 0; index < args.length; index++) {
    if (args[index] === '--write') write = true;
    else if (args[index] === '--project' && args[index + 1] && !args[index + 1].startsWith('--')) {
      directory = args[++index];
    } else throw new Error('Usage: prepare-targets.cjs [--write] [--project WORKSPACE_SUBDIRECTORY]');
  }
  const project = fs.realpathSync(path.resolve(root, directory));
  const relative = path.relative(root, project);
  if (relative === '..' || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new Error('Project must stay inside the workspace.');
  }
  const file = path.join(project, 'package.json');
  const original = fs.readFileSync(file, 'utf8');
  const manifest = JSON.parse(original);
  const edits = [];
  for (const [name, version] of Object.entries(plan.packages)) {
    const sections = ['dependencies', 'devDependencies'].filter(section =>
      Object.hasOwn(manifest[section] || {}, name));
    if (sections.length !== 1) throw new Error(`${name}: expected exactly one direct declaration; inspect the plan.`);
    const section = sections[0];
    const from = manifest[section][name];
    if (from !== version) edits.push({ name, section, from, to: version });
    manifest[section][name] = version;
  }
  const undeclared = Object.keys({ ...manifest.dependencies, ...manifest.devDependencies })
    .filter(name => (name.startsWith('@angular/') || name.startsWith('@angular-devkit/')) && !plan.packages[name]);
  if (undeclared.length) throw new Error(`Additional Angular packages need coordinated review: ${undeclared.join(', ')}`);
  const sha256 = createHash('sha256').update(original).digest('hex');
  if (write) {
    fs.writeFileSync(file, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
  }
  console.log(JSON.stringify({ mode: write ? 'write' : 'dry-run', file, beforeSha256: sha256, edits }, null, 2));
}

try { main(); } catch (error) { console.error(error.message); process.exitCode = 1; }
