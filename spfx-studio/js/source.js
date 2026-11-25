// source.js — the same web part as a standard SPFx 1.21.1 project (npm install, gulp bundle --ship,
// gulp package-solution --ship). Same solution, feature and component ids, so a package built from source
// upgrades the one downloaded from the studio. The runtime and your code are embedded as plain JavaScript.
import { RUNTIME_SOURCE } from './runtime.js';
import { DEPS, SPFX_VERSION, names, runtimeSpec, scopeCss, componentVersion, solutionVersion, manifest } from './bundle.js';

const j = (o) => JSON.stringify(o, null, 2) + '\n';

export function sourceFiles(project) {
  const n = names(project), s = project.spec;
  const folder = n.alias.replace(/WebPart$/, '').replace(/^./, (c) => c.toLowerCase());
  const base = `src/webparts/${folder}/${n.alias}`;
  const m = manifest(project, 'unused.js');
  const man = { $schema: 'https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json', ...m, version: '*' };
  delete man.loaderConfig;
  const files = [];
  files.push(['package.json', j({
    name: n.pkg, version: componentVersion(project), private: true, engines: { node: '>=22.14.0 < 23.0.0' }, main: 'lib/index.js',
    scripts: { build: 'gulp bundle', clean: 'gulp clean', test: 'gulp test', ship: 'gulp clean && gulp bundle --ship && gulp package-solution --ship' },
    dependencies: Object.fromEntries([['tslib', '2.3.1'], ...DEPS.map((d) => [d[0], SPFX_VERSION])]),
    devDependencies: {
      '@microsoft/rush-stack-compiler-5.3': '0.1.0', '@rushstack/eslint-config': '4.0.1', '@microsoft/eslint-plugin-spfx': SPFX_VERSION, '@microsoft/eslint-config-spfx': SPFX_VERSION,
      '@microsoft/sp-build-web': SPFX_VERSION, '@types/webpack-env': '~1.15.2', ajv: '^6.12.5', eslint: '8.57.1', gulp: '4.0.2', typescript: '~5.3.3', '@microsoft/sp-module-interfaces': SPFX_VERSION,
    },
  })]);
  files.push(['config/config.json', j({ $schema: 'https://developer.microsoft.com/json-schemas/spfx-build/config.2.0.schema.json', version: '2.0',
    bundles: { [n.entry]: { components: [{ entrypoint: `./lib/webparts/${folder}/${n.alias}.js`, manifest: `./${base}.manifest.json` }] } }, externals: {}, localizedResources: {} })]);
  files.push(['config/package-solution.json', j({ $schema: 'https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json',
    solution: { name: n.solution, id: project.ids.solution, version: solutionVersion(project), includeClientSideAssets: true, skipFeatureDeployment: s.tenantWide !== false, isDomainIsolated: false,
      developer: { name: s.developer || 'SPFx Studio', websiteUrl: '', privacyUrl: '', termsOfUseUrl: '', mpnId: `Undefined-${SPFX_VERSION}` },
      metadata: { shortDescription: { default: s.description || s.title }, longDescription: { default: s.description || s.title }, screenshotPaths: [], videoUrl: '', categories: [] },
      features: [{ title: `${n.pkg} Feature`, description: `The feature that activates elements of the ${n.pkg} solution.`, id: project.ids.feature, version: '1.0.0.0' }] },
    paths: { zippedPackage: `solution/${n.pkg}.sppkg` } })]);
  files.push(['config/serve.json', j({ $schema: 'https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json', port: 4321, https: true, initialPage: 'https://{tenantDomain}/_layouts/workbench.aspx' })]);
  files.push(['config/write-manifests.json', j({ $schema: 'https://developer.microsoft.com/json-schemas/spfx-build/write-manifests.schema.json', cdnBasePath: '<!-- PATH TO CDN -->' })]);
  files.push(['config/sass.json', j({ $schema: 'https://developer.microsoft.com/json-schemas/core-build/sass.schema.json' })]);
  files.push(['gulpfile.js', `'use strict';\n\nconst build = require('@microsoft/sp-build-web');\n\nvar getTasks = build.rig.getTasks;\nbuild.rig.getTasks = function () {\n  var result = getTasks.call(build.rig);\n  result.set('serve', result.get('serve-deprecated'));\n  return result;\n};\n\nbuild.initialize(require('gulp'));\n`]);
  files.push(['tsconfig.json', j({ extends: './node_modules/@microsoft/rush-stack-compiler-5.3/includes/tsconfig-web.json',
    compilerOptions: { target: 'es2017', forceConsistentCasingInFileNames: true, module: 'esnext', moduleResolution: 'node', jsx: 'react', declaration: true, sourceMap: true, experimentalDecorators: true, skipLibCheck: true, outDir: 'lib', inlineSources: false, noImplicitAny: false,
      typeRoots: ['./node_modules/@types', './node_modules/@microsoft'], types: ['webpack-env'], lib: ['es2017', 'dom', 'dom.iterable'] },
    include: ['src/**/*.ts', 'src/**/*.tsx'] })]);
  files.push(['.eslintrc.js', `require('@rushstack/eslint-config/patch/modern-module-resolution');\nmodule.exports = { extends: ['@microsoft/eslint-config-spfx/lib/profiles/default'], parserOptions: { tsconfigRootDir: __dirname } };\n`]);
  files.push(['.gitignore', 'node_modules\nlib\ndist\ntemp\nrelease\nsharepoint/solution\n*.scss.ts\n']);
  files.push(['src/index.ts', '// A file is required to be in the root of the /src directory by the TypeScript compiler\n']);
  files.push([`${base}.manifest.json`, j(man)]);
  files.push([`${base}.ts`, `/* eslint-disable */
// @ts-nocheck
// ${s.title}: generated by SPFx Studio. The runtime (property pane, list helpers, theme) comes first, then your code (APP).
// Edit APP and CSS, then: npm install, gulp bundle --ship, gulp package-solution --ship.
${DEPS.map((d) => `import * as ${d[2]} from '${d[0]}';`).join('\n')}

const R = ${RUNTIME_SOURCE};

const SPEC = ${JSON.stringify(runtimeSpec(project), null, 2)};

const CSS = ${JSON.stringify(scopeCss(s.css, n.scope))};

const APP = async function (ctx) {
  var h = ctx.h, icon = ctx.icon, esc = ctx.esc, fmt = ctx.fmt, props = ctx.props, sp = ctx.sp;
  {
${String(s.code || '').split('\n').map((l) => '    ' + l).join('\n')}
  }
};

export default R({ ${DEPS.map((d) => d[2]).join(', ')} }, SPEC, CSS, APP);
`]);
  files.push(['README.md', `# ${s.title}\n\n${s.description || ''}\n\nGenerated by SPFx Studio (SPFx ${SPFX_VERSION}, no framework).\n\n## Build\n\n    npm install\n    gulp bundle --ship\n    gulp package-solution --ship\n\nThe package is written to \`sharepoint/solution/${n.pkg}.sppkg\`.\n\n## Lists\n\n${(s.lists || []).map((l) => `- **${l.name}**: ${(l.columns || []).map((c) => `${c.name} (${c.type})`).join(', ')}`).join('\n') || 'None.'}\n\nThe web part's property pane has a button per list that creates it with every column.\n${s.notes ? `\n## Notes\n\n${s.notes}\n` : ''}`]);
  files.push(['studio-project.json', j(project)]);
  return { files, folder: n.pkg };
}
