'use strict';

const path = require('node:path');
const { pathToFileURL } = require('node:url');
const root = path.resolve(__dirname, '../../..');

async function main() {
  const file = process.argv[2];
  if (!['tsconfig.json', 'tsconfig.app.json', 'tsconfig.spec.json'].includes(file)) {
    throw new Error('Choose tsconfig.json, tsconfig.app.json or tsconfig.spec.json.');
  }
  const resolveFromProject = name => require.resolve(name, { paths: [root] });
  const compiler = await import(pathToFileURL(resolveFromProject('@angular/compiler-cli')).href);
  const ts = require(resolveFromProject('typescript'));
  const { options, errors } = compiler.readConfiguration(path.join(root, file));
  console.log(JSON.stringify({ file, options: {
    strict: options.strict,
    strictTemplates: options.strictTemplates,
    strictInjectionParameters: options.strictInjectionParameters,
    strictInputAccessModifiers: options.strictInputAccessModifiers,
    typeCheckHostBindings: options.typeCheckHostBindings,
    moduleResolution: ts.ModuleResolutionKind[options.moduleResolution],
    target: ts.ScriptTarget[options.target], module: ts.ModuleKind[options.module], lib: options.lib
  }, errors: errors.map(error => ts.flattenDiagnosticMessageText(error.messageText, '\n')) }, null, 2));
  if (errors.length) process.exitCode = 1;
}

main().catch(error => { console.error(error.message); process.exitCode = 1; });
