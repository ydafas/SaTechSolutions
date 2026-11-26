// sppkg.js — writes the .sppkg (an Open Packaging Conventions zip) in the browser, matching what
// `gulp package-solution --ship` writes for SPFx 1.21: AppManifest, one feature with the web part's element
// manifest, a ClientSideAssets feature that carries the bundle (includeClientSideAssets: true), the _rels and
// [Content_Types].xml. Tenant-scoped deployment (skipFeatureDeployment) so it can be made available to all sites.
import { buildWebPart, names, solutionVersion, SPFX_VERSION, uuid } from './bundle.js';

const XML = '<?xml version="1.0" encoding="utf-8"?>';
const REL = 'http://schemas.microsoft.com/sharepoint/2012/app/relationships/';
const x = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
const attr = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

const CONTENT_TYPES = XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  [['xml', 'text/xml'], ['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['webpart', 'text/xml'], ['htm', 'text/html'], ['html', 'text/html'], ['aspx', 'text/xml'], ['resx', 'text/xml'], ['js', 'application/javascript'], ['json', 'application/json'], ['png', 'image/png'], ['jpg', 'image/jpeg'], ['bmp', 'image/bmp'], ['gif', 'image/gif']]
    .map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"></Default>`).join('') + '</Types>';

function rels(list, start) {
  return XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    list.map(([type, target], i) => `<Relationship Type="${REL}${type}" Target="${attr(target)}" Id="r${start + i}"></Relationship>`).join('') + '</Relationships>';
}
const partConfig = (id) => XML + `<AppPartConfig xmlns="http://schemas.microsoft.com/sharepoint/2012/app/partconfiguration"><Id>${id}</Id></AppPartConfig>`;
const feature = (title, desc, id) => XML + `<Feature xmlns="http://schemas.microsoft.com/sharepoint/" Title="${attr(title)}" Description="${attr(desc)}" Id="${id}" Version="1.0.0.0" Scope="Web" Hidden="FALSE"></Feature>`;

/** Every file of the package, as [path, string]. Deterministic apart from the two part-config ids. */
export function packageFiles(project, { configIds = [uuid(), uuid()] } = {}) {
  const wp = buildWebPart(project), n = names(project), s = project.spec;
  const featureId = project.ids.feature, ver = solutionVersion(project);
  const dev = JSON.stringify({ name: s.developer || 'SPFx Studio', websiteUrl: '', privacyUrl: '', termsOfUseUrl: '', mpnId: `Undefined-${SPFX_VERSION}` });
  const featureFile = `feature_${featureId}.xml`;
  const elementFile = `${featureId}/WebPart_${project.ids.component}.xml`;
  const files = [];
  files.push(['AppManifest.xml', XML + `<App xmlns="http://schemas.microsoft.com/sharepoint/2012/app/manifest" Name="${attr(n.solution)}" ProductID="${project.ids.solution}" SharePointMinVersion="16.0.0.0" IsClientSideSolution="true" Version="${ver}" SkipFeatureDeployment="${s.tenantWide === false ? 'false' : 'true'}" IsDomainIsolated="false"><Properties><Title>${x(s.title)}</Title><DeveloperProperties>${x(dev)}</DeveloperProperties><ShortDescription><LocalizedString CultureName="default">${x(s.description || s.title)}</LocalizedString></ShortDescription><LongDescription><LocalizedString CultureName="default">${x(s.description || s.title)}</LocalizedString></LongDescription><Screenshots></Screenshots></Properties></App>`]);
  files.push(['_rels/AppManifest.xml.rels', rels([['manifest-feature', '/' + featureFile], ['manifest-clientsideasset', '/ClientSideAssets.xml']], 5)]);
  files.push(['_rels/.rels', rels([['package-manifest', '/AppManifest.xml']], 0)]);
  files.push(['[Content_Types].xml', CONTENT_TYPES]);
  files.push([featureFile, feature(`${n.pkg} Feature`, `The feature that activates elements of the ${n.pkg} solution.`, featureId)]);
  files.push([`_rels/${featureFile}.rels`, rels([['partconfiguration', `/${featureFile}.config.xml`], ['feature-elementmanifest', '/' + elementFile]], 1)]);
  files.push([`${featureFile}.config.xml`, partConfig(configIds[0])]);
  files.push([elementFile, XML + `<Elements xmlns="http://schemas.microsoft.com/sharepoint/"><ClientSideComponent Name="${attr(n.alias.replace(/WebPart$/, ''))}" Id="${project.ids.component}" ComponentManifest="${attr(JSON.stringify(wp.manifest))}" Type="WebPart"></ClientSideComponent><Module Name="${attr(n.alias.replace(/WebPart$/, ''))}" Url="_catalogs/wp" List="113"></Module></Elements>`]);
  files.push(['ClientSideAssets.xml', feature('Client Side Assets', 'A feature that help deploy client side component assets to SharePoint Online.', project.ids.assets)]);
  files.push(['_rels/ClientSideAssets.xml.rels', rels([['partconfiguration', '/ClientSideAssets.xml.config.xml'], ['clientsideasset', `/ClientSideAssets/${wp.file}`]], 3)]);
  files.push(['ClientSideAssets.xml.config.xml', partConfig(configIds[1])]);
  files.push([`ClientSideAssets/${wp.file}`, wp.text]);
  return { files, webPart: wp, fileName: `${n.pkg}.sppkg` };
}

/** Zip it (JSZip must be loaded). Returns {blob|uint8array, fileName}. */
export async function buildSppkg(project, JSZip, type = 'blob') {
  const { files, fileName, webPart } = packageFiles(project);
  const z = new JSZip();
  for (const [p, c] of files) z.file(p, c, { date: new Date() });
  const data = await z.generateAsync({ type, compression: 'DEFLATE', mimeType: 'application/octet-stream' });
  return { data, fileName, webPart };
}
