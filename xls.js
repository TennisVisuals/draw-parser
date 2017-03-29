!function() {

   let tp = function() {};

   /* tables */
   var draw_byes = {
      '12': [1, 4, 9, 12],
      '24': [1, 6, 7, 12, 13, 18, 19, 24],
      '48': [1, 6, 7, 12, 13, 18, 19, 24, 25, 30, 31, 36, 37, 42, 43, 48],
   }
   var main_draw_rounds = ['F', 'SF', 'QF', 'R16', 'R32', 'R64', 'R128'];

   tp.config = {
      score:   /[\d\(]+[\d\.\(\)\[\]\\ \-\,\/O]+(Ret)?(ret)?(RET)?[\.]*$/,
      ended:   ['ret.', 'RET', 'Def.', 'def.', 'BYE', 'w.o', 'W.O', 'Abandoned'],
   }

   tp.points = {};
   tp.profiles = {
      'TP': {
         identification: {
            includes: ['WS', 'WD', 'MS', 'MD'],
            sub_includes: ['BS', 'BD', 'GS', 'GD'],
         },
         columns: {
            position: 'A',
            rank:     'C',
            players:  'E',
            country:  'D',
            rounds:   'F',
         },
         rows: { header:   4, },
         gaps: { draw:     { term: 'Round 1', gap: 0 }, },
         header_columns: [
            { attr: 'id',        header: 'Member ID' },
            { attr: 'entry',     header: 'St.' },
            { attr: 'club',      header: 'Club' },
            { attr: 'country',   header: 'Cnty' },
            { attr: 'rank',      header: 'Rank' },
            { attr: 'players',   header: 'Round 1' },
         ],
         player_rows: { player_names: true },
         extraneous: {},
         routines: {},
      },
   }

   /* parsing */
   let unique = (arr) => arr.filter((item, i, s) => s.lastIndexOf(item) == i);
   let includes = (list, elements) => elements.map(e => list.indexOf(e) >= 0).reduce((a, b) => a || b);
   let subInclude = (list, elements) => list.map(e => includes(e, elements)).reduce((a, b) => a || b);
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
   let getValue = (sheet, reference) => tp.value(sheet[reference]);
   let numberValue = (sheet, reference) => !isNaN(parseInt(getValue(sheet, reference))) ? parseInt(getValue(sheet, reference)) : '';
   let letterValue = (letter) => parseInt(letter, 36) - 9;
   let getRow = (reference) => reference ? parseInt(/\d+/.exec(reference)[0]) : undefined;
   let getCol = (reference) => reference ? reference[0] : undefined;
   let findValueRefs = (search_text, sheet) => Object.keys(sheet).filter(ref => tp.value(sheet[ref]) == search_text);
   let validRanking = (value) => /^\d+$/.test(value) || /^MR\d+$/.test(value);
   let inDrawColumns = (ref, round_columns) => round_columns.indexOf(ref[0]) >= 0;
   let inDrawRows = (ref, range) => getRow(ref) >= +range[0] && getRow(ref) <= +range[1];
   let cellsContaining = ({sheet, term}) => {
      let references = Object.keys(sheet);
      return references.filter(ref => (sheet[ref].v + '').toLowerCase().indexOf(term.toLowerCase()) >= 0);
   }
   let findGaps = ({sheet, term}) => {
      let gaps = [];
      let gap_start = 0;
      let instances = cellsContaining({sheet, term}).map(reference => getRow(reference)).filter((item, i, s) => s.lastIndexOf(item) == i);
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
   let extraneousData = (sheet, ref) => {
      let value = sheet[ref].v;
      if (!isNaN(value) && value < 16) return true;
      let extraneous = tp.profiles[tp.profile].extraneous;
      if (extraneous && extraneous.starts_with) {
         let cell_value = tp.value(sheet[ref]) + '';
         return extraneous.starts_with.map(s => cell_value.toLowerCase().indexOf(s) == 0).reduce((a, b) => (a || b));
      }
   }
   let scoreOrPlayer = ({cell_value, players}) => {
      let draw_position = tp.drawPosition({ full_name: cell_value, players });
      if (draw_position) return true;
      let score = cell_value.match(tp.config.score);
      if (score && score[0] == cell_value) return true;
      let ended = tp.config.ended.map(ending => cell_value.indexOf(ending) >= 0).reduce((a, b) => a || b);
      if (ended) return true;

      if (tp.verbose) console.log('Not Score or Player:', cell_value);
      return false;
   }
   let lastFirstI = (name) => {
      if (name.indexOf(',') >= 0) {
         let components = name.toLowerCase().split(',').map(m=>m.trim());
         return components[1] ? `${components[0]}, ${components[1][0]}` : '';
      }
      let components = name.toLowerCase().split('[')[0].split(' ').filter(f=>f);
      return components.length ? `${components[0][0]}, ${components.reverse()[0]}` : '';
   }
   let headerColumns = ({sheet}) => {
      let profile = tp.profiles[tp.profile];
      let columns = Object.assign({}, profile.columns);
      if (profile.header_columns) {
         profile.header_columns.forEach(obj => {
            let col = getCol(findValueRefs(obj.header, sheet)[0]);
            if (col) columns[obj.attr] = col;
         });
      }
      return columns;
   }
   let normalizeScore = (score) => {
      return score;
   }

   /* exportable functions */
   tp.value = (cell) => {
      let val = cell ? cell.v + '' : '';
      return (typeof val == 'string') ? val.trim() : val;
   }
   tp.playerRows = ({sheet, draw = 'draw'}) => {
      let profile = tp.profiles[tp.profile];
      let columns = headerColumns({sheet});

      let rr_result = [];
      let player_names = Object.keys(sheet)
         .filter(f => f[0] == columns.players && getRow(f) > profile.rows.header)
         .filter(f => tp.value(sheet[f]) && typeof tp.value(sheet[f]) == 'string')
         .map(f=>getRow(f));
      let draw_positions = Object.keys(sheet)
         .filter(f => f[0] == columns.position && /\d/.test(f[1]) && /^\d+(a)?$/.test(tp.value(sheet[f])))
         .map(ref=>getRow(ref));
      let rankings = Object.keys(sheet)
         .filter(f => f[0] == columns.rank && /\d/.test(f[1]) && validRanking(tp.value(sheet[f])))
         .map(ref=>getRow(ref));
      let finals;

      // check whether this is Round Robin
      if (columns.rr_result) {
         rr_result = Object.keys(sheet)
            .filter(f => f[0] == columns.rr_result && /\d/.test(f[1]) && /^\d+[\.]*$/.test(tp.value(sheet[f])))
            .map(ref=>getRow(ref));
         rankings = rankings.filter(f => rr_result.indexOf(f) >= 0);
      }

      let sources = [draw_positions, rankings, rr_result];

      // Necessary for finding all player rows in TP Doubles Draws
      if (profile.player_rows && profile.player_rows.player_names) {
         let additions = [];
         player_names.forEach(f => {
            // additions is just a counter
            if (tp.value(sheet[`${columns.players}${f}`]).toLowerCase() == 'bye') additions.push(f - 1); 
         });
         sources.push(player_names);
      }

      let rows = [].concat(...sources).filter((item, i, s) => s.lastIndexOf(item) == i).sort((a, b) => a - b);

      if (profile.gaps && profile.gaps.draw) {
         let gaps = findGaps({sheet, term: profile.gaps[draw].term}); 
         if (gaps.length) {
            let gap = gaps[profile.gaps[draw].gap];
            if (!columns.rr_result) {
               // filter rows by gaps unless Round Robin Results Column
               rows = rows.filter(row => row > gap[0] && row < gap[1]);
            } else {
               // names that are within gap in round robin
               finals = player_names.filter(row => row > gap[0] && row < gap[1]);
            }
         }
      }
      let range = [rows[0], rows[rows.length - 1]];

      // determine whether there are player rows outside of Round Robins
      finals = finals ? finals.filter(f => rows.indexOf(f) < 0) : undefined;
      finals = finals && finals.length ? finals : undefined;

      return { rows, range, finals };
   }
   tp.roundColumns = ({sheet}) => {
      let header_row = tp.profiles[tp.profile].rows.header;
      let rounds_column = tp.profiles[tp.profile].columns.rounds;
      let columns = Object.keys(sheet)
         .filter(key => key.length == 2 && key.slice(1) == header_row && letterValue(key[0]) >= letterValue(rounds_column))
         .map(m=>m[0]).filter((item, i, s) => s.lastIndexOf(item) == i).sort();
      return columns;
   }
   tp.roundData = ({sheet, player_data, round_robin}) => {
      let players = player_data.players;
      let round_columns = tp.roundColumns({sheet});
      let range = player_data.range;
      let cell_references = Object.keys(sheet)
         .filter(ref => inDrawColumns(ref, round_columns) && inDrawRows(ref, range))
         .filter(ref => !extraneousData(sheet, ref));

      let filtered_columns = round_columns.map(column => { 
         let column_references = cell_references.filter(ref => ref[0] == column).filter(ref => scoreOrPlayer({ cell_value: tp.value(sheet[ref]), players }));
         return { column, column_references, }
      }).filter(f=>f.column_references.length);

      // work around for round robins with blank 'BYE' columns
      let start = round_columns.indexOf(filtered_columns[0].column);
      let end = round_columns.indexOf(filtered_columns[filtered_columns.length - 1].column);
      let column_range = round_columns.slice(start, end);
      let rr_columns = column_range.map(column => { 
         let column_references = cell_references.filter(ref => ref[0] == column).filter(ref => scoreOrPlayer({ cell_value: tp.value(sheet[ref]), players }));
         return { column, column_references, }
      });

      return round_robin ? rr_columns : filtered_columns;
   }
   tp.drawPlayers = ({sheet}) => {
      let extract_seed = /\[(\d+)(\/\d+)?\]/;
      let columns = headerColumns({sheet});
      let rows, range;
      ({rows, range, finals} = tp.playerRows({sheet}));
      let players = rows.map(row => {
         let draw_position = numberValue(sheet, `${columns.position}${row}`);

         // MUST BE DOUBLES
         if (!draw_position) draw_position = numberValue(sheet, `${columns.position}${row + 1}`);

         let player = { draw_position };
         if (columns.seed) player.seed = numberValue(sheet, `${columns.seed}${row}`);

         let full_name = getValue(sheet, `${columns.players}${row}`);
         if (extract_seed.test(full_name)) {
            player.seed = parseInt(extract_seed.exec(full_name)[1]);
            full_name = full_name.split('[')[0].trim();
         }

         player.full_name = full_name;
         player.last_first_i = lastFirstI(full_name);
         if (columns.id) player.id = getValue(sheet, `${columns.id}${row}`);
         if (columns.club) player.club = getValue(sheet, `${columns.club}${row}`);
         if (columns.rank) player.rank = numberValue(sheet, `${columns.rank}${row}`);
         if (columns.entry) player.entry = getValue(sheet, `${columns.entry}${row}`);
         if (columns.country) player.country = getValue(sheet, `${columns.country}${row}`);
         if (columns.rr_result) player.rr_result = numberValue(sheet, `${columns.rr_result}${row}`);
         return player;
      });
      return { players, rows, range, finals };
   }

   tp.drawPosition = ({full_name, players, idx = 0}) => {
      // idx used in instances where there are multiple BYEs, such that they each have a unique draw_position
      let tournament_player = players.filter(player => player.full_name && player.full_name == full_name)[idx];
      if (!tournament_player) {
         // find player draw position by last name, first initial; for draws where first name omitted after first round
         tournament_player = players.filter(player => player.last_first_i && player.last_first_i == lastFirstI(full_name))[0];
      }
      return tournament_player ? tournament_player.draw_position : undefined;
   }

   let columnMatches = (sheet, round, players) => {
      let names = [];
      let matches = [];
      let winners = [];
      let last_draw_position;
      let round_occurrences = [];
      let last_row_number = getRow(round.column_references[0]) - 1;
      round.column_references.forEach((reference, index) => {
         // if row number not sequential => new match
         let this_row_number = getRow(reference);
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

         let cell_value = tp.value(sheet[reference]);
         let idx = names.filter(f => f == cell_value).length;
         // names used to keep track of duplicates, i.e. 'BYE' such that
         // a unique draw_position is returned for subsequent byes
         names.push(cell_value);
         let draw_position = tp.drawPosition({ full_name: cell_value, players, idx });

         // cell_value is a draw position => round winner(s)
         if (draw_position != undefined) {
            last_draw_position = draw_position;
            if (winners.indexOf(draw_position) < 0) winners.push(draw_position);
         } else {
            // cell_value is not draw position => match score
            if (last_draw_position) {
               // keep track of how many times draw position occurs in column
               if (!round_occurrences[last_draw_position]) round_occurrences[last_draw_position] = [];
               round_occurrences[last_draw_position].push(matches.length);
            }
            matches.push({ winners, result: normalizeScore(cell_value) });
            winners = [];
         }
      });
      // still winners => last column match had a bye
      if (winners.length) matches.push({ bye: winners });
      round_occurrences = round_occurrences.map((indices, draw_position) => { return { draw_position, indices }}).filter(f=>f);
      return { round_occurrences, matches };
   }

   let addByes = (rounds, players) => {
      let profile = tp.profiles[tp.profile];
      if (draw_byes[players.length] && profile.routines && profile.routines.add_byes) {
         let round_winners = [].concat(...rounds[0].map(match => match.winners).filter(f=>f));
         draw_byes[players.length].forEach(player => { 
            if (round_winners.indexOf(player) < 0) rounds[0].push({ bye: [player] }); 
         });
         rounds[0].sort((a, b) => {
            let adp = a.winners ? a.winners[0] : a.bye[0];
            let bdp = b.winners ? b.winners[0] : b.bye[0];
            return adp - bdp;
         });
      } else {
         draw_byes[players.length] = [];
      }
      return rounds;
   }

   let add1stRound = (rounds, players) => {
      // 1st round players are players without byes or wins 
      let winners = unique([].concat(...rounds.map(matches => [].concat(...matches.map(match => match.winners).filter(f=>f)))));
      let notWinner = (draw_position) => winners.indexOf(draw_position) < 0;
      let notBye = (draw_position) => !draw_byes[players.length] || draw_byes[players.length].indexOf(draw_position) < 0;
      let first_round_losers = players
         .filter(player => notWinner(player.draw_position) && notBye(player.draw_position))
         .map(m=>m.draw_position)
         .filter((item, i, s) => s.lastIndexOf(item) == i)
         .map(m => { return { players: [m] }});
      rounds.push(first_round_losers);
      return rounds;
   }

   let findEmbeddedRounds = (rounds) => {
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
      return embedded_rounds;
   }

   let constructMatches = (rounds, players) => {
      let draw_type = rounds[0].length == 1 ? 'main' : 'qualification';
      rounds.forEach((round, index) => {
         if (index + 2 == rounds.length) round = round.filter(player => player.bye == undefined);
         if (index + 1 < rounds.length) {
            let round_matches = [];
            let round_winners = [];
            round.forEach(match => {
               let player = match.winners ? match.winners[0] : match.bye ? match.bye[0] : match.players[0];
               round_winners.push(player);
               round_matches.push(match);
            });
            let previous_round_players = rounds[index + 1].map(match => {
               return match.winners ? match.winners[0] : match.bye ? match.bye[0] : match.players[0];
            });
            let eliminated_players = previous_round_players.filter(player => round_winners.indexOf(player) < 0);
            let draw_positions = players.map(m=>m.draw_position).filter((item, i, s) => s.lastIndexOf(item) == i).length;
            let round_name = index + 2 < rounds.length || index < 3 ? main_draw_rounds[index] : `R${draw_positions}`;
            round_matches.forEach((match, match_index) => {
               match.round = draw_type == 'main' ? round_name : `Q${index || ''}`;
               match.losers = [eliminated_players[match_index]];
               match.loser_names = players.filter(f=>f.draw_position == eliminated_players[match_index]).map(p=>p.full_name);
            });
         }
      });
      return rounds;
   }

   let findPlayerAtDrawPosition = (players, start, goal, direction) => {
      let index = start + direction;
      while (players[index] && players[index].draw_position != goal && index < players.length && index >= 0) { index += direction; }
      if (!players[index]) return undefined;
      return index;
   }

   let determineWinner = (score) => {
      let tally = [0, 0];
      let set_scores = score.split(' ');
      set_scores.forEach(set_score => {
         if (/\d+[\(\)\-\/]*/.test(set_score)) tally[parseInt(set_score[0]) > parseInt(set_score[1]) ? 0 : 1] += 1
      });
      if (tally[0] > tally[1]) return 0;
      if (tally[1] > tally[0]) return 1;
      return undefined;
   }

   let reverseScore = (score) => {
      return score.split(' ').map(set_score => {
         let tiebreak = /\((\d+)\)/.exec(set_score);
         let score = set_score.split('(')[0];
         let scores = (/\d+/.test(score)) ? score.split('').reverse().join('') : score;
         if (tiebreak) scores += `${tiebreak[0]}`;
         return scores;
      }).join(' ');
   }

   tp.tournamentDraw = ({sheet, player_data}) => {
      let rounds = [];
      let matches = [];
      player_data = player_data || tp.drawPlayers({sheet});
      let players = player_data.players;
      let round_robin = players.length ? players.map(p=>p.rr_result != undefined).reduce((a, b) => a || b) : false;

      if (round_robin) {
         let hash = [];
         let player_rows = player_data.rows;
         let pi = player_data.players.map((p, i) => p.rr_result ? i : undefined).filter(f=>f != undefined);
         let group_size = pi.length;

         // combine all cell references that are in result columns
         let round_data = tp.roundData({sheet, player_data, round_robin: true});
         let rr_columns = round_data.map(m=>m.column).slice(0, group_size);
         let result_references = [].concat(...round_data.map((round, index) => index < group_size ? round.column_references : []));
         player_rows.forEach((player_row, player_index) => {
            let player_result_referencess = result_references.filter(ref => ref.slice(1) == player_row);
            player_result_referencess.forEach(reference => {
               let result_column = reference[0];
               let player_draw_position = players[player_index].draw_position;
               let opponent_draw_position = rr_columns.indexOf(result_column) + 1;
               let direction = opponent_draw_position > player_draw_position ? 1 : -1;
               let opponent_index = findPlayerAtDrawPosition(players, player_index, opponent_draw_position, direction);
               let result = normalizeScore(tp.value(sheet[reference]));
               let match_winner = determineWinner(result);
               let loser = match_winner ? player_index : opponent_index;
               let winner = match_winner ? opponent_index : player_index;

               if (players[loser] && players[winner] && match_winner != undefined) {
                  if (match_winner) result = reverseScore(result);
                  let match = { 
                     winner_names: [players[winner].full_name],
                     loser_names: [players[loser].full_name],
                     round: 'RR' + players[winner].rr_result,
                     result,
                  };

                  // don't add the same match twice
                  if (hash.indexOf(`${winner}${loser}${result}`) < 0) {
                     hash.push(`${winner}${loser}${result}`);
                     matches.push(match);
                  }
               }
            });
         });

         // also search for final match in sheet
         let profile = tp.profiles[tp.profile];
         if (player_data.finals && profile.targets && profile.targets.winner) {
            let columns = headerColumns({sheet});
            let keys = Object.keys(sheet);
            let target = unique(keys.filter(f=>sheet[f].v == profile.targets.winner))[0];
            if (target && target.match(/\d+/)) {
               let finals_col = target[0];
               let finals_row = parseInt(target.match(/\d+/)[0]);
               let finals_range = player_data.finals.filter(f => f != finals_row);
               let finals_cells = keys.filter(k => {
                  let numeric = k.match(/\d+/);
                  if (!numeric) return false;
                  return numeric[0] >= finals_range[0] && numeric[0] <= finals_range[finals_range.length - 1] && k[0] == finals_col;
               }).filter(ref => scoreOrPlayer({ cell_value: tp.value(sheet[ref]), players }));
               let finals_details = finals_cells.map(fc => sheet[fc].v);
               let finalists = player_data.finals
                  .map(row => getValue(sheet, `${columns.players}${row}`))
                  .filter(player => scoreOrPlayer({ cell_value: player, players }));
               let winner = finals_details.filter(f => finalists.indexOf(f) >= 0)[0];
               let result = finals_details.filter(f => finalists.indexOf(f) < 0)[0];
               let loser = finalists.filter(f => f != winner)[0];
               let match = {
                  winner_names: [winner],
                  loser_names: [loser],
                  round: 'F',
                  result,
               }
               matches.push(match);
            }
         }

      } else {
         let first_round;
         let round_data = tp.roundData({sheet, player_data});
         rounds = round_data.map(round => {
            let column_matches = columnMatches(sheet, round, players);
            let matches_with_results = column_matches.matches.filter(match => match.result);
            if (!matches_with_results.length) {
               // first_round necessary for situations where i.e. 32 draw used when 16 would suffice
               first_round = column_matches.matches.filter(match => match.winners).map(match => match.winners[0]);
            }
            return column_matches;
         });
         findEmbeddedRounds(rounds).forEach(round => rounds.push(round));
         rounds = rounds.map(round => round.matches);
         if (!rounds.length) {
            if (tp.verbose) console.log('ERROR WITH SHEET - Possibly abandoned.', tp.profile, 'format.');
            return { rounds, matches: [] };
         }
         rounds = addByes(rounds, players);
         /* reverse rounds to:
            - append first round to end
            - start identifying matches with Final
            - filter players with byes into 2nd round
         */
         rounds.reverse();

         if (first_round) {
            let filtered_players = players.filter(player => first_round.indexOf(player.draw_position) >= 0);
            rounds = add1stRound(rounds, filtered_players);
         } else {
            rounds = add1stRound(rounds, players);
         }
         rounds = rounds.filter(round => round.filter(f=> { return f.winners ? f.winners.length : true }).length);
         rounds = constructMatches(rounds, players);

         // merge all rounds into list of matches
         matches = [].concat(...rounds).filter(f=>f.losers && f.result);

         // add player names to matches
         matches.forEach(match => match.winner_names = players.filter(f=>f.draw_position == match.winners[0]).map(p=>p.full_name));
      }

      return { rounds, matches };
   }

   tp.calcPoints = (category = 20, tournament_rank, round, format, draw_positions, score, profile) => {
      profile = profile ? tp.profiles[profile] : tp.profiles[tp.profile];
      if (!profile.points) return undefined;
      let multiplier = category > 12 ? Math.pow(2, (category - 12) / 2) : 1;
      // draw_positions is total # of draw positions
      round = round != 'QF' && round.indexOf('Q') >= 0 ? `${draw_positions}${round}` : round;
      let points_row = tp.points[profile.points][format][round];
      let points = points_row && points_row[tournament_rank - 1] ? points_row[tournament_rank - 1] * multiplier : 0;
      if (score.indexOf('w.o.') >= 0) return 0;
      return points;
   }

   // matches must be sorted with finals first
   tp.playerPoints = (matches, category, rank, profile) => {
      let player_points = { singles: {}, doubles: {} };
      matches.forEach(match => {
         let format = match.winner_2 ? 'doubles' : 'singles';
         let points = tp.calcPoints(category || match.tournament_category, rank || match.tournament_rank, match.round, format, match.draw_positions, match.score, profile);
         let pp = player_points[format];
         if (match.draw_type != 'consolation') {
         }
         if (points && !pp[match.winner_1] || points > pp[match.winner_1]) {
            player_points[format][match.winner_1] = points;
            if (match.winner_2) player_points[format][match.winner_2] = points;
         }
      });
      return player_points;
   }

   tp.drawResults = (workbook) => {
      let rows = [];
      tp.setWorkbookProfile({workbook});

      let draw_type;
      let tournament_rank;
      let tournament_category;
      let tournament_data = [];
      tournament_data = tp.profile == 'HTS' ? tp.HTS_tournamentData(workbook) : {};
      if (Object.keys(tournament_data).length) {
         tournament_rank = parseInt(tournament_data.rang_turnira.match(/\d+/)[0]);
         tournament_category = tournament_data.kategorija.match(/\d+/);
         tournament_category = tournament_category ? parseInt(tournament_category[0]) : 'Senior';
      }

      let processDraw = (sheet_name) => {
         let sheet = workbook.Sheets[sheet_name];
         let player_data = tp.drawPlayers({sheet});
         let players = player_data.players;
         let draw = tp.tournamentDraw({sheet, player_data});
         let playerData = (name) => players.filter(player => player.full_name == name)[0];
         let draw_positions = players.map(m=>m.draw_position).filter((item, i, s) => s.lastIndexOf(item) == i).length;
         let round_robin = players.length ? players.map(p=>p.rr_result != undefined).reduce((a, b) => a || b) : false;

         if (tp.profile == 'TP') {
            let number = /\d+/;
            let type = tp.value(sheet['A2']);
            tournament_category = number.test(type) ? number.exec(type) : undefined;
         }

         // TODO: draw_type
         // HTS UTJEŠNI TURNIR == Consolation => no points
         
         draw.matches.forEach(match => {
            let points = 0;
            let format = match.winner_names.length == 2 ? 'doubles' : 'singles';
            let pd = {
               w1: playerData(match.winner_names[0]),
               w2: playerData(match.winner_names[1]),
               l1: playerData(match.loser_names[0]),
               l2: playerData(match.loser_names[1]),
            }
            let row = {};
            if (tp.profile == 'HTS') {
               Object.assign(row, {
                  id: tournament_data.id_turnira,
                  tournament_date: tournament_data.datum_turnir,
                  tournament_rank,
                  tournament_category,
               });
            }
            Object.assign(row, {
               round: match.round,
               score: match.result,
               winner_1: match.winner_names[0],
               winner_1_rank: pd.w1 ? pd.w1.rank : '',
               winner_2: match.winner_names[1] || '',
               winner_2_rank: pd.w2 ? pd.w2.rank : '',
               loser_1: match.loser_names[0], 
               loser_1_rank: pd.l1 ? pd.l1.rank : '',
               loser_2: match.loser_names[1] || '',
               loser_2_rank: pd.l2 ? pd.l2.rank : '',
               draw_positions,
            });
            if (match.winner_names[0]) rows.push(row);
         });
      }

      workbook.SheetNames.filter(sheet_name => {
         /* attempting to include sheet filter declarations in profiles
         let include = true;
         Object.keys(tp.profiles)
            .map(profile => profile.sheet_filter && exclude(sheet_filter, sheet_name) ? false : true)
            .reduce((a, b) => a || b);
         */
         if (tp.profile == 'HTS') {
            if (sheet_name.toLowerCase().indexOf('raspored') >= 0) return false;
//            if (sheet_name.toLowerCase().indexOf('rr') >= 0) return false;
            return sheet_name.match(/\d+/) || sheet_name.match(/_M/) || sheet_name.match(/_Z/);
         }
         return sheet_name;
      }).forEach(sheet_name => {
         if (tp.verbose) console.log('processing draw:', sheet_name);
         processDraw(sheet_name);
      });
      return { rows };
   }

   tp.setWorkbookProfile = ({workbook}) => {
      let sheet_names = workbook.SheetNames;

      Object.keys(tp.profiles)
         .forEach(profile => {
            let identification = tp.profiles[profile].identification;
            if (identification.includes && includes(sheet_names, identification.includes)) tp.profile = profile;
            if (identification.sub_includes && subInclude(sheet_names, identification.sub_includes)) tp.profile = profile;
         });
   }

   if (typeof define === "function" && define.amd) define(tp); else if (typeof module === "object" && module.exports) module.exports = tp;
   this.tp = tp;
 
}();
