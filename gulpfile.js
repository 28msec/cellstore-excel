'use strict';

var gulp = require('gulp');
var $ = require('gulp-load-plugins')();
var minimist = require('minimist');

var isOnTravis = process.env.CIRCLECI === 'true';
var artifactsDir = process.env.CIRCLE_ARTIFACTS;
var isOnTravisAndMaster = isOnTravis && process.env.CI_PULL_REQUEST === '' && process.env.CIRCLE_BRANCH === 'master';

var config = {
    ExcelDnaVersion: '0.33.9',
    CellStoreVersion: '1.0.0',
    NewtonsoftVersion: '7.0.1',
    RestSharp: '105.1.0'
};
var isWindows = /^win/.test(process.platform);
var nugetCmd = isWindows ? 'nuget' : 'mono nuget.exe';
var compileCmd = isWindows ? 'csc' : 'mcs -sdk:4.5';
var packCmd = isWindows ? 'ExcelDnaPack' : 'mono ExcelDnaPack.exe';

var args = minimist(process.argv.slice(2),
  {
    alias: { d: 'debug' },
    default: { debug: false },
    '--': true
  });

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
    'cd build && mkdir bin32 && mkdir bin64 && mkdir release'
]));

gulp.task('install-dependencies', ['clean'], $.shell.task([
    isWindows ? ':' : 'cd build && wget https://nuget.org/nuget.exe',
    'cd build && ' + nugetCmd + ' install CellStore.NET -Version ' + config.CellStoreVersion,

    // install Excel DNA
    'cd build && ' + nugetCmd + ' install ExcelDna.AddIn -Version ' + config.ExcelDnaVersion,
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDnaPack.exe build/bin32/ExcelDnaPack.exe',
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDna.Integration.dll build/bin32/ExcelDna.Integration.dll',
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDna.xll build/bin32/CellStore.Excel.xll',

    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDnaPack.exe build/bin64/ExcelDnaPack.exe',
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDna.Integration.dll build/bin64/ExcelDna.Integration.dll',
    'cp build/ExcelDna.AddIn.' + config.ExcelDnaVersion + '/tools/ExcelDna64.xll build/bin64/CellStore.Excel.xll',

    // install Cellstore
    'cp build/CellStore.NET.' + config.CellStoreVersion + '/lib/CellStore.dll build/bin32/CellStore.dll',
    'cp build/CellStore.NET.' + config.CellStoreVersion + '/lib/CellStore.dll build/bin64/CellStore.dll',
    'cp lib/CellStore.Excel.dna build/bin32/CellStore.Excel.dna',
    'cp lib/CellStore.Excel.dna build/bin64/CellStore.Excel.dna',

    // install Newtonsoft
    'cp build/Newtonsoft.Json.' + config.NewtonsoftVersion + '/lib/net45/Newtonsoft.Json.dll build/bin32/Newtonsoft.Json.dll',
    'cp build/Newtonsoft.Json.' + config.NewtonsoftVersion + '/lib/net45/Newtonsoft.Json.dll build/bin64/Newtonsoft.Json.dll',

    // install RestSharp
    'cp build/RestSharp.' + config.RestSharp + '/lib/net45/RestSharp.dll build/bin32/RestSharp.dll',
    'cp build/RestSharp.' + config.RestSharp + '/lib/net45/RestSharp.dll build/bin64/RestSharp.dll'

]));

gulp.task('compile', $.shell.task([
    pathFix(compileCmd + ' -r:build/bin32/ExcelDna.Integration.dll,build/bin32/Newtonsoft.Json.dll,build/bin32/RestSharp.dll,build/bin32/CellStore.dll,System.Windows.Forms.dll,System.Runtime.Caching.dll -target:library -out:build/bin32/CellStore.Excel.dll -recurse:src/*.cs -platform:x86' + (args.debug ? ' -debug' : '')),
    pathFix(compileCmd + ' -r:build/bin64/ExcelDna.Integration.dll,build/bin64/Newtonsoft.Json.dll,build/bin64/RestSharp.dll,build/bin64/CellStore.dll,System.Windows.Forms.dll,System.Runtime.Caching.dll -target:library -out:build/bin64/CellStore.Excel.dll -recurse:src/*.cs -platform:x64' + (args.debug ? ' -debug' : ''))
]));

gulp.task('build', ['compile'], $.shell.task([
    'cd build && cd bin32 && ' + packCmd + ' CellStore.Excel.dna /Y',
    'cd build && cd bin64 && ' + packCmd + ' CellStore.Excel.dna /Y',

    'cp build/bin32/CellStore.Excel-packed.xll build/release/CellStore.Excel.x86.32.xll',
    'cp build/bin64/CellStore.Excel-packed.xll build/release/CellStore.Excel.x64.64.xll'
]));

gulp.task('artifacts', $.shell.task([
    artifactsDir ? 'cd build && cp -R * ' + artifactsDir : 'echo "CIRCLE_ARTIFACTS not set"'
]));

gulp.task('publish', function(){
    gulp.src(['./build/release/CellStore.Excel.x86.32.xll','./build/release/CellStore.Excel.x64.64.xll'])
        .pipe($.githubRelease({
            repo: 'cellstore-excel',
            owner: '28msec',
            manifest: require('./package.json')
        }));
});

gulp.task('release', function(done){
    if(isOnTravis) {
        // @TODO We cannot automatically release because ExcelDnaPack doesn't work on linux
        $.runSequence('install-dependencies', 'compile', 'artifacts', function () {
            //$.runSequence('publish', done);
            done();
        });
    } else {
        $.runSequence('install-dependencies', 'build', done);
    }
});

gulp.task('default', ['build']);
