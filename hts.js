!function() {

   function hts() { }

   /* local */
   let fs = require('fs-extra');
   hts.tp = require('./xls');
   hts.xlsx = require('xlsx');
   hts.cache = './excel/';

   let loadFile = (dir, filename) => fs.readFileSync('./' + path.join(hts.cache, dir, filename), 'utf8');
   let cacheFile = (data, filename, dir, format = 'utf8') => {
      if (!filename) return;
      let target_dir = './' + path.join(hts.cache + dir) + '/';
      fs.ensureDirSync(target_dir);
      fs.writeFileSync(target_dir + filename, data, format);
   }
   let fileList = ({dir = '/', require, exclude} = {}) => {
      let target_dir = path.join(hts.cache, dir);
      fs.ensureDirSync(target_dir);
      var files = fs.readdirSync(target_dir).reduce(function(list, file) {
        var name = path.join(target_dir, file);
        var isDir = fs.statSync(name).isDirectory();
        return list.concat(isDir ? fileList({dir: `${dir}/${file}/`, require, exclude}) : [{ dir: dir, file: file }]);
      }, []);
      files = files.filter(({file}) => file[0] != '.');
      if (require) { files = files.filter(({file}) => file.indexOf(require) >= 0); }
      if (exclude) { files = files.filter(({file}) => file.indexOf(exclude) < 0); }
      return files;
   }

   hts.workbookList = ({require, exclude} = {}) => fileList({require, exclude}).map(workbook => './' + path.join(hts.cache, workbook.dir, workbook.file));
   hts.loadWorkbook = (file) => {
      let workbook = hts.xlsx.readFile(file);
      hts.tp.setWorkbookProfile({workbook});
      return workbook;
   }

   hts.saveObject = ({json, name, dir = 'Processed'}) => cacheFile(JSON.stringify(json), `${name}.json`, dir);
   hts.loadObject = ({name, dir = 'Processed'}) => JSON.parse(loadFile(dir, `${name}.json`));

   hts.allWorkbooks = (workbook_files) => {
      let rows = [];
      hts.tp.verbose = true;
      workbook_files.forEach((workbook_file, index) => {
         console.log('processing:', workbook_file, 'index:', index);
         let workbook = hts.loadWorkbook(workbook_file);
         let results = hts.tp.drawResults(workbook);
         rows = rows.concat(...results.rows);
      });
      return rows;
   }

   hts.allTournaments = (workbooks = hts.workbookList()) => workbooks.map(workbook => hts.tp.HTS_tournamentData(hts.loadWorkbook(workbook)));
   /* end local */

   /* browser */
   /*
   hts.loadWorkbook = (file_content) => {
      let workbook = XLSX.read(file_content, {type: 'binary'});
      tp.setWorkbookProfile({workbook});
      return workbook;
   }
   */

   tp.points.HTS = {
      singles: {
         'F':    [150,120,90,75,60,40,30,12],   '32Q':  [16,14,12,10,8,6,4,0],
         'SF':   [100,75,60,48,36,28,20,8],     '32Q1': [8,7,6,5,4,3,2,0],
         'QF':   [75,60,46,32,26,20,16,4],      '32Q2': [4,4,3,2,2,1,1,0],
         'R12':  [50,40,32,26,20,16,12,0],      '32Q3': [2,2,0,0,0,0,0,0],
         'R16':  [50,40,32,26,20,16,12,0],      '48Q':  [8,7,6,5,4,3,2,0],
         'R24':  [32,28,24,20,16,12,8,0],       '48Q1': [4,4,3,2,2,1,1,0],
         'R32':  [32,28,24,20,16,12,8,0],       '48Q2': [2,2,0,0,0,0,0,0],
         'R48':  [16,14,12,10,8,6,4,0],         '64Q':  [8,7,6,5,4,3,2,0],
         'R64':  [16,14,12,10,8,6,4,0],         '64Q1': [4,4,3,2,2,1,1,0],
                                                '64Q2': [2,2,0,0,0,0,0,0],

         'RR1':  [0,0,0,0,0,0,0,12],            'RRQ':  [0,14,12,10,8,6,4,0],
         'RR2':  [0,0,0,0,0,0,0,8],             'RRQ1': [0,7,6,5,4,3,2,0],
         'RR3':  [0,0,0,0,0,0,0,4],             'RRQ2': [0,4,3,2,2,1,1,0],
         'RR4':  [0,0,0,0,0,0,0,2],
      },
      doubles: {
         'F':    [38,30,23,19,15,10,8],
         'SF':   [25,19,15,12,9,7,5],
         'QF':   [19,15,12,8,6,5,4],	
         'R12':  [12,10,8,6,5,4,3],
         'R16':  [12,10,8,6,5,4,3],
      },
   };

   tp.profiles.HTS = {
      identification: { includes: ['Pocetna'], },
      sheeet_filter: { includes: ['raspored', 'rr'] },
      points: 'HTS',
      columns: {
         position: 'A',
         rank:     'B',
         entry:    'C',
         seed:     'D',
         players:  'E',
         club:     'G',
         rounds:   'H',
      },
      rows: { header:   6, },
      gaps: { 
         draw:     { term: 'rang', gap: 0 }, 
         preround: { term: 'rang', gap: 1 },
      },
      header_columns: [
         { attr: 'rr_result',     header: 'Poredak' },
      ],
      targets: { winner: 'Pobjednik', },
      extraneous: { starts_with: ['gs', 'bez', 'pobjed', 'final'], },
      routines: { add_byes:   true, }
   };

   tp.HTS_tournamentData = (workbook) => {
      if (workbook.SheetNames.indexOf('Pocetna') < 0) return {};
      let details = workbook.Sheets['Pocetna'];
      return {
         naziv_turnira:    tp.value(details.B7),
         kategorija:       tp.value(details.B10),
         sve_kategorije:   tp.value(details.C10),
         datum_turnir:     tp.value(details.B12),
         datu_rang_liste:  tp.value(details.B14),
         rang_turnira:     tp.value(details.B16),
         id_turnira:       tp.value(details.C12),
         organizator:      tp.value(details.D12),
         mjesto:           tp.value(details.E12),
         vrhovni_sudac:    tp.value(details.F12),
         director_turnira: tp.value(details.E14),
         dezurni_ljiecnik: tp.value(details.F14),
         singlovi:         workbook.SheetNames.filter(f=>f.indexOf('Si ') == 0),
         dubl:             workbook.SheetNames.filter(f=>f.indexOf('Do ') == 0),
      };
   }

   if (typeof define === "function" && define.amd) define(hts); else if (typeof module === "object" && module.exports) module.exports = hts;
   this.hts = hts;
 
}();
