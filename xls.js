/* parser for HTS legacy Excel files */

module.exports = function() {

   function tp() { }
   let fs = require('fs-extra');
   let removeDiacritics = require('diacritics').remove;
   tp.xlsx = require('xlsx');
   tp.cache = './excel/';
   tp.profile = 'HTS';

   tp.not = [];

   /* tables */
   var draw_byes = {
      '12': [1, 4, 9, 12],
      '24': [1, 6, 7, 12, 13, 18, 19, 24],
      '48': [1, 6, 7, 12, 13, 18, 19, 24, 25, 30, 31, 36, 37, 42, 43, 48],
   }
   var main_draw_rounds = ['F', 'SF', 'QF', 'R16', 'R32', 'R64', 'R128'];
   tp.extraneous = { starts_with: ['gs', 'bez', 'pobjed', 'final'], }

   tp.points = {
      'HTS': {
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
      },
   }

   tp.profiles = {
      'HTS': {
         columns: {
            position: 'A',
            rank:     'B',
            entry:    'C',
            seed:     'D',
            players:  'E',
            club:     'G',
         },
         gaps: { 
            draw:     { term: 'rang', gap: 0 }, 
            preround: { term: 'rang', gap: 1 },
         },
      }
   }

   /* analysis */
   tp.gatherSheetNames = (require = 'xlsm') => [].concat(...tp.workbookList({require}).map(workbook => tp.xlsx.readFile(workbook, { bookSheets: true, }).SheetNames));
   tp.nameCount = (sheet_names) => {
      count = {};
      sheet_names.forEach(name => { if (Object.keys(count).indexOf(name) >= 0) { count[name] += 1; } else { count[name] = 1; } });
      return count;
   }

   /* utilities */
   tp.saveObject = ({json, name, dir = 'Processed'}) => tp.cacheFile(JSON.stringify(json), `${name}.json`, dir);
   tp.loadObject = ({name, dir = 'Processed'}) => JSON.parse(tp.loadFile(dir, `${name}.json`));
   tp.loadFile = (dir, filename) => fs.readFileSync('./' + path.join(tp.cache, dir, filename), 'utf8');
   tp.cacheFile = (data, filename, dir, format = 'utf8') => {
      if (!filename) return;
      let target_dir = './' + path.join(tp.cache + dir) + '/';
      fs.ensureDirSync(target_dir);
      fs.writeFileSync(target_dir + filename, data, format);
   }

   tp.fileList = fileList;
   function fileList({dir = '/', require, exclude} = {}) {
      let target_dir = path.join(tp.cache, dir);
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
   let findMiddle = (arr) => arr[Math.round((arr.length - 1) / 2)];
   let findMiddles = (arr, number) => {
      if (!(arr.length % 2)) return [];
      let parts = [arr.slice()];
      let middles;
      while (number) {
         middles = [];
         let more_parts = [];
         parts.forEach(part => {
            let middle = findMiddle(part);
            middles.push(middle);
            more_parts.push(part.slice(0, middle));
            more_parts.push(part.slice(middle + 1));
            parts = more_parts;
         });
         number--;
      }
      return middles;
   }

   /* parsing */
   tp.workbookList = ({require} = {}) => tp.fileList({require}).map(workbook => './' + path.join(tp.cache, workbook.dir, workbook.file));
   tp.loadWorkbook = (file) => tp.xlsx.readFile(file);
   tp.row = (cell) => parseInt(/\d+/.exec(cell)[0]);
   tp.cellsContaining = ({sheet, term}) => {
      let cells = Object.keys(sheet);
      return cells.filter(cell => (sheet[cell].v + '').toLowerCase().indexOf(term.toLowerCase()) >= 0);
   }
   tp.findGaps = ({sheet, term}) => {
      let gaps = [];
      let gap_start = 0;
      let instances = tp.cellsContaining({sheet, term}).map(cell => tp.row(cell)).filter((item, i, s) => s.lastIndexOf(item) == i);
      instances.unshift(0);
      let nextGap = (index) => { 
         while (instances[index + 1] == instances[index] + 1 && index < instances.length) { index += 1; }
         return index;
      }
      let gap_end = nextGap(0);
      while (gap_end < instances.length) {
         if (gap_start) gaps.push([instances[gap_start], instances[gap_end]]);
         gap_start = nextGap(gap_end); 
         gap_end = nextGap(gap_start + 1);
      }
      return gaps;
   }
   tp.validRanking = (value) => /^\d+$/.test(value) || /^MR\d+$/.test(value);
   tp.playerRows = ({sheet, draw = 'draw'}) => {
      // for doubles the first player row is before the first draw position row
      let profile = tp.profiles[tp.profile];
      let columns = profile.columns;
      let draw_positions = Object.keys(sheet).filter(f=>f[0] == columns.position && /\d/.test(f[1]) && /^\d+$/.test(sheet[f].v)).map(m=>tp.row(m));
      let rankings = Object.keys(sheet).filter(f=>f[0] == columns.rank && /\d/.test(f[1]) && tp.validRanking(sheet[f].v)).map(m=>tp.row(m));
      let rows = [].concat(draw_positions, rankings).filter((item, i, s) => s.lastIndexOf(item) == i).sort((a, b) => a - b);
      if (profile.gaps && profile.gaps.draw) {
         let gaps = tp.findGaps({sheet, term: profile.gaps[draw].term}); 
         let gap = gaps[profile.gaps[draw].gap];
         rows = rows.filter(row => row > gap[0] && row < gap[1]);
      }
      let range = [rows[0], rows[rows.length - 1]];

      return { rows, draw_positions, range };
   }
   tp.columnCells = (sheet, column) => Object.keys(sheet).filter(key=>key[0] == column);
   tp.inDrawColumns = (cell, columns) => columns.indexOf(cell[0]) >= 0;
   tp.inDrawRows = (cell, range) => tp.row(cell) >= +range[0] && tp.row(cell) <= +range[1];

   // TODO: modify to support ITF Juniors Draw Format
   tp.roundColumns = ({sheet}) => Object.keys(sheet).filter(key => key.length == 2 && key.slice(1) == '6' && 'ABCDEFG'.split('').indexOf(key[0]) < 0).map(m=>m[0]);
   tp.extraneousData = (sheet, cell) => !isNaN(sheet[cell].v) || tp.extraneous.starts_with.map(s => sheet[cell].v.toLowerCase().indexOf(s) == 0).reduce((a, b) => (a || b));
   tp.scoreOrPlayer = ({value, tournament_players}) => {
      let draw_position = tp.drawPosition({ full_name: value, tournament_players });
      if (draw_position) return true;
      let score = value.trim().match(/[\d\(]+[\d\(\\ \-\,\/)O]+(Ret)?(ret)?(RET)?[\.]*$/);
      if (score && score[0] == value.trim()) return true;

      let ended = ['ret.', 'RET', 'def.', 'BYE', 'w.o', 'W.O'].map(ending => value.indexOf(ending) >= 0).reduce((a, b) => a || b);
      if (ended) return true;

      console.log('Not Score or Player:', value);
      tp.not.push(`|${value}|`);
      return false;
   }
   tp.roundData = ({sheet, tournament_players}) => {
      let columns = tp.roundColumns({sheet});
      let range = tp.playerRows({sheet}).range;
      let cells = Object.keys(sheet)
         .filter(cell => tp.inDrawColumns(cell, columns) && tp.inDrawRows(cell, range))
         .filter(cell => !tp.extraneousData(sheet, cell));
      let result = columns.map(column => { 
         let name = sheet[`${column}6`].v;
         let column_cells = cells.filter(cell => cell[0] == column).filter(cell => tp.scoreOrPlayer({ value: sheet[cell].v, tournament_players }));
         return { column, name, cells: column_cells, }
      }).filter(f=>f.cells.length);
      return result;
   }
   tp.tournamentData = (workbook) => {
      if (workbook.SheetNames.indexOf('Pocetna') < 0) return {};
      let details = workbook.Sheets['Pocetna'];
      return {
         naziv_turnira: details.B7 ? details.B7.v : '',
         kategorija: details.B10 ? details.B10.v : '',
         sve_kategorije: details.C10 ? details.C10.v : '',
         datum_turnir: details.B12 ? details.B12.v : '',
         datu_rang_liste: details.B14 ? details.B14.v : '',
         rang_turnira: details.B15 ? details.B16.v : '',
         id_turnira: details.C12 ? details.C12.v : '',
         organizator: details.D12 ? details.D12.v : '',
         mjesto: details.E12 ? details.E12.v : '',
         vrhovni_sudac: details.F12 ? details.F12.v : '',
         director_turnira: details.E14 ? details.E14.v : '',
         dezurni_ljiecnik: details.F14 ? details.F14.v : '',
         singlovi: workbook.SheetNames.filter(f=>f.indexOf('Si ') == 0),
         dubl: workbook.SheetNames.filter(f=>f.indexOf('Do ') == 0),
      };
   }
   tp.lastFirstI = (name) => {
      let components = name.toLowerCase().split(',').map(m=>removeDiacritics(m.trim()));
      return components[1] ? `${components[0]}, ${components[1][0]}` : '';
   }
   tp.drawPlayers = ({sheet}) => {
      let columns = tp.profiles[tp.profile].columns;
      let rows = tp.playerRows({sheet}).rows;
      return rows.map(row => {
         let full_name = sheet[`${columns.players}${row}`] ? sheet[`${columns.players}${row}`].v : '';
         let last_first_i = tp.lastFirstI(full_name);
         let draw_position = parseInt(sheet[`A${row}`] ? sheet[`A${row}`].v : '');
         let player = { full_name, last_first_i, draw_position };
         if (columns.seed) player.seed = sheet[`${columns.seed}${row}`] ? sheet[`${columns.seed}${row}`].v : '';
         if (columns.club) player.club = sheet[`${columns.club}${row}`] ? sheet[`${columns.club}${row}`].v : '';
         if (columns.rank) player.rank = sheet[`${columns.rank}${row}`] ? sheet[`${columns.rank}${row}`].v : '';
         if (columns.entry) player.entry = sheet[`${columns.entry}${row}`] ? sheet[`${columns.entry}${row}`].v : '';
         return player;
      });
   }

   // find player draw position by last name (without diacritics), first initial; for draws where first name omitted after first round
   tp.drawPosition = ({full_name, tournament_players}) => {
      let tournament_player = tournament_players.filter(player => player.last_first_i && player.last_first_i == tp.lastFirstI(full_name))[0];
      return tournament_player ? tournament_player.draw_position : undefined;
   }
   tp.separateRounds = ({sheet}) => {
      let all_round_winners = [];
      let tournament_players = tp.drawPlayers({sheet});
      let round_data = tp.roundData({sheet, tournament_players});

      let columnMatches = (round) => {
         let matches = [];
         let round_occurrences = [];

         // accumulate winners for doubles
         let winners = [];

         let last_draw_position;
         let last_row_number = tp.row(round.cells[0]) - 1;
         round.cells.forEach((cell, index) => {
            // if row number not sequential => new match
            let this_row_number = tp.row(round.cells[index]);
            if (this_row_number != last_row_number + 1 && winners.length) {
               if (last_draw_position) {
                  // keep track of how many times draw position occurs in column
                  if (!round_occurrences[last_draw_position]) round_occurrences[last_draw_position] = [];
                  round_occurrences[last_draw_position].push(matches.length);
               }
               matches.push({ winners, });
               winners = [];
            }
            last_row_number = this_row_number;

            let value = sheet[cell].v.trim();
            let draw_position = tp.drawPosition({ full_name: value, tournament_players });
            // value is a draw position => round winner(s)
            if (draw_position != undefined) {
               last_draw_position = draw_position;
               if (winners.indexOf(draw_position) < 0) winners.push(draw_position);
               if (all_round_winners.indexOf(draw_position) < 0) all_round_winners.push(draw_position);
            } else {
               // value is not draw position => match score
               if (last_draw_position) {
                  // keep track of how many times draw position occurs in column
                  if (!round_occurrences[last_draw_position]) round_occurrences[last_draw_position] = [];
                  round_occurrences[last_draw_position].push(matches.length);
               }
               matches.push({ winners, result: value });
               winners = [];
            }
         });
         // still winners => last column match had a bye
         if (winners.length) matches.push({ bye: winners });
         round_occurrences = round_occurrences.map((indices, draw_position) => { return { draw_position, indices }}).filter(f=>f);
         return { round_occurrences, matches };
      }
      let rounds = round_data.map(round => columnMatches(round));

      // check each round for embedded rounds
      let embedded_rounds = [];
      rounds.forEach((round, index) => {
         let embedded = round.round_occurrences.filter(f=>f.indices.length > 1).length;
         if (embedded) {
            let other_rounds = [];
            let indices = [...Array(round.matches.length)].map((_, i) => i);
            for (let i=embedded; i > 0; i--) {
               let embedded_indices = findMiddles(indices, i);
               if (embedded_indices.length) {
                  other_rounds = other_rounds.concat(...embedded_indices);
                  embedded_rounds.push({ matches: embedded_indices.map(match_index => Object.assign({}, round.matches[match_index])) });
                  embedded_indices.forEach(match_index => round.matches[match_index].result = undefined);
               }
            }
            // filter out embedded matches
            round.matches = round.matches.filter(match => match.result);
         }
      });
      embedded_rounds.forEach(round => rounds.push(round));
      rounds = rounds.map(round => round.matches);

      if (!rounds.length) {
         console.log('ERROR WITH FILE');
         return { rounds, matches: [] };
      }

      // add seeded players with byes to 2nd round players
      // rounds[0] is 2nd round before 1st round players added below
      // TODO: currently only works for singles draws
      if (draw_byes[tournament_players.length]) {
         draw_byes[tournament_players.length].forEach(player => { rounds[0].push({ bye: [player] }); });
         rounds[0].sort((a, b) => {
            let adp = a.winners ? a.winners[0] : a.bye[0];
            let bdp = b.winners ? b.winners[0] : b.bye[0];
            return adp - bdp;
         });
      } else {
         draw_byes[tournament_players.length] = [];
      }

      /* reverse rounds to:
         - append first round to end
         - start identifying matches with Final
         - filter players with byes into 2nd round
      */
      rounds.reverse();

      // 1st round players are tournament_players without byes or wins 
      let notWinner = (draw_position) => all_round_winners.indexOf(draw_position) < 0;
      let notBye = (draw_position) => draw_byes[tournament_players.length].indexOf(draw_position) < 0;
      let first_round_losers = tournament_players
         .filter(player => notWinner(player.draw_position) && notBye(player.draw_position))
         .map(m=>m.draw_position)
         .filter((item, i, s) => s.lastIndexOf(item) == i)
         .map(m => { return { players: [m] }});
      rounds.push(first_round_losers);

      let draw_type = rounds[0].length == 1 ? 'main' : 'qualification';
      rounds.forEach((round, index) => {
         if (index + 2 == rounds.length) round = round.filter(player => player.bye == undefined);
         if (index + 1 < rounds.length) {
            let round_matches = [];
            let round_winners = [];
            // let previous_round_players = [];
            round.forEach(match => {
               let player = match.winners ? match.winners[0] : match.bye ? match.bye[0] : match.players[0];
               round_winners.push(player);
               round_matches.push(match);
            });
            let previous_round_players = rounds[index + 1].map(match => {
               return match.winners ? match.winners[0] : match.bye ? match.bye[0] : match.players[0];
            });
            /*
            rounds[index + 1].forEach(match => {
               let player = match.winners ? match.winners[0] : match.bye ? match.bye[0] : match.players[0];
               previous_round_players.push(player);
            });
            */
            let eliminated_players = previous_round_players.filter(player => round_winners.indexOf(player) < 0);
            /*
            console.log('round', index);
            console.log('round matches', round_matches);
            console.log('round winners', round_winners.length);
            console.log('previous round players', previous_round_players.length);
            console.log('eliminated players', eliminated_players.length);
            */
            let draw_positions = tournament_players.map(m=>m.draw_position).filter((item, i, s) => s.lastIndexOf(item) == i).length;
            let round_name = index + 2 < rounds.length || index < 3 ? main_draw_rounds[index] : `R${draw_positions}`;
            round_matches.forEach((match, match_index) => {
               match.round = draw_type == 'main' ? round_name : `Q${index || ''}`;
               match.losers = [eliminated_players[match_index]];
               match.loser_names = tournament_players.filter(f=>f.draw_position == eliminated_players[match_index]).map(p=>p.full_name);
            });
         }
      });

      // merge all rounds into list of matches
      let matches = [].concat(...rounds).filter(f=>f.losers && f.result);

      // add player names to matches
      matches.forEach(match => match.winner_names = tournament_players.filter(f=>f.draw_position == match.winners[0]).map(p=>p.full_name));
      return { rounds, matches };
   }
   tp.calcPoints = (category = 20, tournament_rank, round, format, draw_positions) => {
      let multiplier = category > 12 ? Math.pow(2, (category - 12) / 2) : 1;
      round = round.indexOf('Q') >= 0 ? `${draw_positions}${round}` : round;
      let points_row = tp.points[tp.profile][format][round];
      return points_row && points_row[tournament_rank - 1] ? points_row[tournament_rank - 1] * multiplier : 0;
   }
   tp.drawResults = (workbook) => {
      let rows = [];
      let player_points = { singles: {}, doubles: {} };
      let tournament_data = tp.tournamentData(workbook);

      if (!Object.keys(tournament_data).length) {
         console.log('NOT AN HTS SPREADSHEET');
         return { rows: [], player_points: [] };
      }

      let tournament_rank = parseInt(tournament_data.rang_turnira.match(/\d+/)[0]);
      let tournament_category = tournament_data.kategorija.match(/\d+/);
      tournament_category = tournament_category ? parseInt(tournament_category[0]) : 'Senior';

      let processDraw = (sheet_name) => {
         let sheet = workbook.Sheets[sheet_name];
         let draw = tp.separateRounds({sheet});
         let players = tp.drawPlayers({sheet});
         let playerData = (name) => players.filter(player => player.full_name == name)[0];
         let draw_positions = players.map(m=>m.draw_position).filter((item, i, s) => s.lastIndexOf(item) == i).length;
         draw.matches.forEach(match => {
            let points = 0;
            let format = match.winner_names.length == 2 ? 'doubles' : 'singles';
            if (!player_points[format][match.winner_names[0]]) {
               points = tp.calcPoints(tournament_category, tournament_rank, match.round, format, draw_positions);
               player_points[format][match.winner_names[0]] = points;
            }
            let pd = {
               w1: playerData(match.winner_names[0]),
               w2: playerData(match.winner_names[1]),
               l1: playerData(match.loser_names[0]),
               l2: playerData(match.loser_names[1]),
            }
            let row = {
               id: tournament_data.id_turnira,
               tournament_date: tournament_data.datum_turnir,
               tournament_rank,
               tournament_category,
               round: match.round,
               score: match.result,
               winner_points: points,
               winner_1: match.winner_names[0],
               winner_1_rank: pd.w1 ? pd.w1.rank : '',
               winner_2: match.winner_names[1] || '',
               winner_2_rank: pd.w2 ? pd.w2.rank : '',
               loser_1: match.loser_names[0], 
               loser_1_rank: pd.l1 ? pd.l1.rank : '',
               loser_2: match.loser_names[1] || '',
               loser_2_rank: pd.l2 ? pd.l2.rank : '',
            }
            rows.push(row);
         });
      }

      workbook.SheetNames.filter(sheet_name => {
         if (sheet_name.toLowerCase().indexOf('raspored') >= 0) return false;
         if (sheet_name.toLowerCase().indexOf('rr') >= 0) return false;
         return sheet_name.match(/\d+/);
      }).forEach(sheet_name => {
         console.log('processing draw:', sheet_name);
         processDraw(sheet_name);
      });
      return { rows, player_points };
   }

   /* routines */
   tp.allWorkbooks = (workbook_files) => {
      let rows = [];
      workbook_files.forEach((workbook_file, index) => {
         console.log('processing:', workbook_file, 'index:', index);
         let workbook = x.loadWorkbook(workbook_file);
         let results = x.drawResults(workbook);
         rows = rows.concat(...results.rows);
      });
      return rows;
   }

   tp.allTournaments = (workbooks = tp.workbookList()) => workbooks.map(workbook => tp.tournamentData(tp.loadWorkbook(workbook)));

   return tp;
}
