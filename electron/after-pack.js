// electron-builder afterPack hook.
// Some transitive npm packages get placed only as nested children of
// their immediate parent (e.g. call-bind-apply-helpers ends up at
// node_modules/call-bind/node_modules/call-bind-apply-helpers/) even
// though dunder-proto and other siblings need to resolve them from
// the top of the tree. We copy each known-missing module up to the
// top level so Node's standard resolution algorithm can find it.

const fs = require('fs');
const path = require('path');

// Recursively copy a directory, mirroring `cp -R`. Avoids pulling in
// a dep just for this — fs.cpSync is in Node 16.7+ which Electron 31
// ships with.
function copyDir(src, dest) {
  fs.cpSync(src, dest, { recursive: true });
}

// Resolve the first directory under any nested node_modules that
// matches `pkgName`. Used as the source when the top-level copy is
// absent.
function findNestedPkg(rootNodeModules, pkgName) {
  const queue = [rootNodeModules];
  while (queue.length) {
    const dir = queue.shift();
    let entries;
    try { entries = fs.readdirSync(dir, { withFileTypes: true }); }
    catch { continue; }
    for (const e of entries) {
      if (!e.isDirectory()) continue;
      const full = path.join(dir, e.name);
      if (e.name === pkgName && fs.existsSync(path.join(full, 'package.json'))) {
        return full;
      }
      if (e.name === 'node_modules') queue.push(full);
      else if (e.name.startsWith('@')) queue.push(full);
      else {
        const nm = path.join(full, 'node_modules');
        if (fs.existsSync(nm)) queue.push(nm);
      }
    }
  }
  return null;
}

// Packages that electron-builder routinely misplaces. Add to this
// list if a future runtime "Cannot find module 'X'" appears in a
// fresh build — copying the nested copy up to the top level is the
// generic remedy.
const FIX_LIST = [
  'call-bind-apply-helpers',
  'es-set-tostringtag',
  'side-channel-list',
  'side-channel-map',
  'side-channel-weakmap',
];

module.exports = async function afterPack(context) {
  const { appOutDir, packager } = context;
  // Resources directory layout differs by platform.
  const isMac = packager.platform.name === 'mac';
  const resourcesDir = isMac
    ? path.join(appOutDir, `${packager.appInfo.productFilename}.app`, 'Contents', 'Resources')
    : path.join(appOutDir, 'resources');
  const appDir = path.join(resourcesDir, 'app');
  const nm     = path.join(appDir, 'node_modules');
  if (!fs.existsSync(nm)) return;  // asar build — nothing to fix here

  for (const pkg of FIX_LIST) {
    const topLevel = path.join(nm, pkg);
    if (fs.existsSync(topLevel)) continue;
    const nested = findNestedPkg(nm, pkg);
    if (!nested) {
      console.warn(`[afterPack] couldn't locate ${pkg} anywhere — leaving as-is`);
      continue;
    }
    copyDir(nested, topLevel);
    console.log(`[afterPack] hoisted ${pkg} → ${path.relative(appOutDir, topLevel)}`);
  }
};
