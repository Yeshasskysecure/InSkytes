'use strict';

const build = require('@microsoft/sp-build-web');
const { ServeTask } = require('@microsoft/gulp-core-build-serve/lib/ServeTask');

ServeTask.prototype._enableCorsMiddleware = function (req, res, next) {
  const requestPath = String(req.url || '').split('?')[0];
  const isScriptOrManifestRequest =
    /\.js$/i.test(requestPath) ||
    requestPath.endsWith('/temp/build/manifests.js') ||
    requestPath.endsWith('/temp/manifests.js');

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Private-Network', 'true');

  if (req.method === 'OPTIONS') {
    res.statusCode = 204;
    res.end();
    return;
  }

  if (isScriptOrManifestRequest) {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    res.removeHeader('ETag');
  }

  next();
};

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.lintCmd.enabled = false;

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  if (typeof result.delete === 'function') {
    result.delete('lint');
  } else {
    delete result.lint;
  }

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
