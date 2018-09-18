// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
// 

var gulp = require('gulp');
var ts = require('gulp-typescript');
var tslint = require('gulp-tslint');
var del = require('del');
var server = require('gulp-develop-server');
var mocha = require('gulp-spawn-mocha');
var sourcemaps = require('gulp-sourcemaps');
var zip = require('gulp-zip');
var browserify = require('browserify');
var source = require('vinyl-source-stream');
var path = require('path');
var minimist = require('minimist');
var fs = require('fs');
var _ = require('lodash');

var knownOptions = {
	string: 'packageName',
	string: 'packagePath',
	string: 'specFilter',
	default: {packageName: 'Package.zip', packagePath: path.join(__dirname, '_package'), specFilter: '*'}
};
var options = minimist(process.argv.slice(2), knownOptions);

var tsProject = ts.createProject('./tsconfig.json', {
    // Point to the specific typescript package we pull in, not a machine-installed one
    typescript: require('typescript'),
});

var filesToWatch = ['**/*.ts', '!node_modules/**'];
var filesToLint = ['**/*.ts', '!src/typings/**', '!node_modules/**'];
var staticFiles = ['src/**/*.json', 'src/**/*.pug', '!src/manifest.json'];
var clientJS = 'build/src/TaskModuleTab.js';
var bundledJS = 'bundle.js';
var msTeamsLib = './node_modules/@microsoft/teams-js/dist/MicrosoftTeams.min.js';

/**
 * Clean build output.
 */
gulp.task('clean', function() {
    return del([
        'build/**/*',
        // Azure doesn't like it when we delete build/src
        '!build/src'
        // 'manifest/**/*'
    ])
});

/**
 * Lint all TypeScript files.
 */
gulp.task('ts:lint', [], function () {
    if (!process.env.GLITCH_NO_LINT) {
        return gulp
            .src(filesToLint)
            .pipe(tslint({
                formatter: 'verbose'
            }))
            .pipe(tslint.report({
                summarizeFailureOutput: true
            }));
      }
});

/**
 * Compile TypeScript and include references to library.
 */
gulp.task('ts', ['clean'], function() {
    return tsProject
        .src()
        .pipe(sourcemaps.init())
        .pipe(tsProject())
        .pipe(sourcemaps.write('.', { sourceRoot: function(file) { return file.cwd + '/build'; }}))
        .pipe(gulp.dest('build/src'));
});

/**
 * Copy statics to build directory.
 */
gulp.task('statics:copy', ['clean'], function () {
    return gulp.src(staticFiles, { base: '.' })
        .pipe(gulp.dest('./build'));
});

/**
 * Copy (generated) client TypeScript files to the /scripts directory
 */
gulp.task('client-js', ['ts'], function() {
    var bundler = browserify({
        entries: clientJS,
        ignoreMissing: true,
        debug: false
    });

    var bundle = function() {
        return bundler
            .bundle()
            .on('.error', function() {})
            .pipe(source(bundledJS))
            .pipe(gulp.dest('./public/scripts'));
    };

    if (global.isWatching) {
        bundler = watchify(bundler);
        bundler.on('update', bundle)
    }

    return bundle();
});

/**
 * Build application.
 */
gulp.task('build', ['clean', 'ts:lint', 'ts', 'client-js', 'statics:copy']);

/**
 * Build manifest
 */
gulp.task('generate-manifest', function() {
    gulp.src(['./public/images/*_icon.png', 'src/manifest.json'])
        .pipe(zip('TaskModule.zip'))
        .pipe(gulp.dest('manifest'));
});

/**
 * Build debug version of the manifest - 
 */
gulp.task('generate-manifest-debug', function() {
    gulp.src(['./public/images/*_icon.png', 'manifest/debug/manifest.json'])
        .pipe(zip('TaskModuleDebug.zip'))
        .pipe(gulp.dest('manifest/debug'));
});

/**
 * Run tests.
 */
gulp.task('test', ['ts', 'statics:copy'], function() {
    return gulp
        .src('build/test/' + options.specFilter + '.spec.js', {read: false})
        .pipe(mocha({cwd: 'build/src'}))
        .once('error', function () {
            process.exit(1);
        })
        .once('end', function () {
            process.exit();
        });
});

/**
 * Package up app into a ZIP file for Azure deployment.
 */
gulp.task('package', ['build'], function () {
    var packagePaths = [
        'build/**/*',
        'public/**/*',
        'web.config',
        'package.json',
        '**/node_modules/**',
        '!build/src/**/*.js.map', 
        '!build/test/**/*', 
        '!build/test', 
        '!build/src/typings/**/*'];

    //add exclusion patterns for all dev dependencies
    var packageJSON = JSON.parse(fs.readFileSync(path.join(__dirname, 'package.json'), 'utf8'));
    var devDeps = packageJSON.devDependencies;
    for (var propName in devDeps) {
        var excludePattern1 = '!**/node_modules/' + propName + '/**';
        var excludePattern2 = '!**/node_modules/' + propName;
        packagePaths.push(excludePattern1);
        packagePaths.push(excludePattern2);
    }

    return gulp.src(packagePaths, { base: '.' })
        .pipe(zip(options.packageName))
        .pipe(gulp.dest(options.packagePath));
});

gulp.task('server:start', ['build'], function() {
    server.listen({path: 'build/src/app.js'}, function(error) {
        console.log(error);
    });
});

gulp.task('server:restart', ['build'], function() {
    server.restart();
});

gulp.task('default', ['server:start'], function() {
    gulp.watch(filesToWatch, ['server:restart']);
});
gulp.task('default', ['clean', 'generate-manifest'], function() {
    console.log('Build completed. Output in manifest folder');
});