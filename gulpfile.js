'use strict';

var gulp = require('gulp');
var $ = require('gulp-load-plugins')();

var isOnTravis = process.env.CIRCLECI === 'true';
var artifactsDir = process.env.CIRCLE_ARTIFACTS;
var isOnTravisAndMaster = isOnTravis && process.env.CI_PULL_REQUEST === '' && process.env.CIRCLE_BRANCH === 'master';

var config = {
    ExcelDnaVersion: '0.33.9',
    CellStoreVersion: '0.0.13',
    NewtonsoftVersion: '7.0.1',
    RestSharp: '105.1.0'
};
var isWindows = /^win/.test(process.platform);
var nugetCmd = isWindows ? 'nuget' : 'mono nuget.exe';
var compileCmd = isWindows ? 'csc' : 'mcs -sdk:4.5';
var packCmd = isWindows ? 'ExcelDnaPack' : 'mono ExcelDnaPack.exe';

var downloadCmd = function(url, output){
    if(isWindows){
        return 'curl -o "' + output + '" ' + url;
    } else {
        return 'wget -O "' + output + '" ' + url;
    }
};

var pathFix = function(str){
    if(isWindows){
        return str.replace(new RegExp('/', 'g'), '\\');
    } else {
        return str;
    }
};

gulp.task('clean', $.shell.task([
    'rm -rf build',
    'mkdir build',
    'cd build && mkdir bin && mkdir release'
]));

gulp.task('install-dependencies', ['clean'], $.shell.task([
    isWindows ? ':' : 'cd build && wget https://nuget.org/nuget.exe',
    'cd build && ' + nugetCmd + ' install ExcelDna.AddIn -Version ' + config.ExcelDnaVersion,
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDnaPack.exe build/bin/ExcelDnaPack.exe',
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDna.Integration.dll build/bin/ExcelDna.Integration.dll',
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDna64.xll build/bin/CellStore.Excel.xll',
    'cd build && ' + nugetCmd + ' install CellStore.NET -Version ' + config.CellStoreVersion,
    'cp lib/CellStore64.Excel.dna build/bin/CellStore.Excel.dna',
    'cp build/CellStore.NET.' + config.CellStoreVersion + '/lib/CellStore.dll build/bin/CellStore.dll',
    'cp build/Newtonsoft.Json.' + config.NewtonsoftVersion + '/lib/net45/Newtonsoft.Json.dll build/bin/Newtonsoft.Json.dll',
    'cp build/RestSharp.' + config.RestSharp + '/lib/net45/RestSharp.dll build/bin/RestSharp.dll'
]));

gulp.task('compile', $.shell.task([
    pathFix(compileCmd + ' -r:build/bin/ExcelDna.Integration.dll,build/bin/Newtonsoft.Json.dll,build/bin/RestSharp.dll,build/bin/CellStore.dll,System.Windows.Forms.dll -target:library -out:build/bin/CellStore.Excel.dll -recurse:src/*.cs -platform:anycpu')
]));

gulp.task('build', ['compile'], $.shell.task([
    'cd build && cd bin && ' + packCmd + ' CellStore.Excel.dna /Y',
    'cp build/bin/CellStore.Excel-packed.xll build/release/CellStore.Excel.xll'
]));

gulp.task('artifacts', $.shell.task([
    artifactsDir ? 'cd build && cp -R * ' + artifactsDir : 'echo "CIRCLE_ARTIFACTS not set"'
]));

gulp.task('publish', function(){
    gulp.src('./build/release/CellStore.Excel.xll')
        .pipe($.githubRelease({
            repo: 'cellstore-excel',
            owner: '28msec',
            manifest: require('./package.json')
        }));
});

gulp.task('release', function(done){
    if(isOnTravisAndMaster) {
        // @TODO ExcelDnaPack doesn't work on linux
        $.runSequence('install-dependencies', 'compile', 'artifacts', function () {
            //$.runSequence('publish', done);
            done();
        });
    } else {
        $.runSequence('install-dependencies', 'build', done);
    }
});

gulp.task('default', ['build']);
