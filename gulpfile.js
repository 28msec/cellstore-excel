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
var nugetCmd = /^win/.test(process.platform) ? 'nuget' : 'mono nuget.exe';
var compileCmd = /^win/.test(process.platform) ? 'csc' : 'mcs';
var packCmd = /^win/.test(process.platform) ? 'ExcelDnaPack' : 'mono ExcelDnaPack.exe';

var downloadCmd = function(url, output){
    if(/^win/.test(process.platform)){
        return 'curl -o "' + output + '" ' + url;
    } else {
        return 'wget -O "' + output + '" ' + url;
    }
};

gulp.task('clean', $.shell.task([
    'rm -rf build',
    'mkdir build',
    'cd build && mkdir bin'
]));

gulp.task('install-dependencies', ['clean'], $.shell.task([
    /^win/.test(process.platform) ? 'ls' : 'cd build && wget https://nuget.org/nuget.exe',
    'cd build && ' + nugetCmd + ' install ExcelDna.AddIn -Version ' + config.ExcelDnaVersion,
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDnaPack.exe build/ExcelDnaPack.exe',
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDna.Integration.dll build/bin/ExcelDna.Integration.dll',
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDna64.xll build/bin/CellStore64.Excel.xll',
    'cd build && ' + nugetCmd + ' install CellStore.NET -Version ' + config.CellStoreVersion,
    'cp lib/CellStore64.Excel.dna build/bin/CellStore64.Excel.dna',
    'cp build/CellStore.NET.' + config.CellStoreVersion + '/lib/CellStore.dll build/bin/CellStore.dll',
    'cp build/Newtonsoft.Json.' + config.NewtonsoftVersion + '/lib/net45/Newtonsoft.Json.dll build/bin/Newtonsoft.Json.dll',
    'cp build/RestSharp.' + config.RestSharp + '/lib/net45/RestSharp.dll build/bin/RestSharp.dll'
]));

gulp.task('build', ['install-dependencies'], $.shell.task([
    compileCmd + ' -sdk:4.5 -r:build/bin/ExcelDna.Integration.dll,build/bin/Newtonsoft.Json.dll,build/bin/RestSharp.dll,build/bin/CellStore.dll -target:library -out:build/bin/CellStore64.Excel.dll -recurse:src/*.cs -platform:anycpu'
]));

gulp.task('pack', ['build'], $.shell.task([
    'cd build && ' + packCmd + ' bin/CellStore64.Excel.dna /Y'
]));

gulp.task('artifacts', $.shell.task([
    'cd build && if [ "' + artifactsDir + '" != "" ] ; then cp -R * ' + artifactsDir + ' ; fi'
]));

gulp.task('release', function(done){
    $.runSequence('pack', 'artifacts', function(){
        if(isOnTravisAndMaster) {
            //$.runSequence('publish', done);
            done();
        } else {
            done();
        }
    });
});

gulp.task('default', ['release']);
