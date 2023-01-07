function getPremierLeagueMembers(): void {

  // シートを取得
  const sheet = SpreadsheetApp.getActiveSheet();

  // データ取得
  const options = {
    method: 'GET',
    headers: {
      'X-RapidAPI-Key': 'c59931cf36msh4636011e33b747ap1d8cadjsn5e2b8c53eadc',
      'X-RapidAPI-Host': 'api-football-v1.p.rapidapi.com'
    }
  };

  const response = UrlFetchApp.fetch('https://api-football-v1.p.rapidapi.com/v3/players?league=39&season=2022', options)
  const json = JSON.parse(response.getContentText());
  const totalPage:number = json.paging.total;
  let members = json.response;
  for (let i = 2; i <= totalPage; i++) {
    Utilities.sleep(2000);
    const response = UrlFetchApp.fetch(`https://api-football-v1.p.rapidapi.com/v3/players?league=39&season=2022&page=${i}`, options)
    const json = JSON.parse(response.getContentText());
    members = members.concat(json.response);
  }

  // データをシートに書き込み
  let row = 2;
  for (let member of members) {
    const { name, age, nationality, height, weight, photo } = member.player;
    const { team, games, substitutes, shots, goals, passes, tackles, duels, dribbles, fouls, cards, penalty } = member.statistics[0];
    const { name: teamName, logo: teamLogo } = team;
    const { appearences, minutes, position, lineups, rating } = games;
    const { in: substitutesIn, out: substitutesOut, bench: substitutesBench } = substitutes;
    const { total: shotsTotal, on: shotsOn } = shots;
    const { total: goalsTotal, conceded: goalsConceded, assists: goalsAssists, saves: goalsSaves } = goals;
    const { total: passesTotal, key: passesKey, accuracy: passesAccuracy } = passes;
    const { total: tacklesTotal } = tackles;
    const { total: duelsTotal, won: duelsWon } = duels;
    const { attempts: dribblesAttempts, success: dribblesSuccess } = dribbles;
    const { drawn: foulsDrawn, commited: foulsCommited } = fouls;
    const { yellow: cardsYellow, yellowred: cardsYellowRed, red: cardsRed } = cards;
    const { scored: penaltyScored, missed: penaltyMissed, saved: penaltySaved } = penalty;

    const data = [
      photo,
      name,
      age,
      position,
      nationality,
      height,
      weight,
      rating,
      teamName,
      teamLogo,
      appearences,
      minutes,
      lineups,
      substitutesIn,
      substitutesOut,
      substitutesBench,
      shotsTotal,
      shotsOn,
      goalsTotal,
      goalsConceded,
      goalsAssists,
      goalsSaves,
      passesTotal,
      passesKey,
      passesAccuracy,
      tacklesTotal,
      duelsTotal,
      duelsWon,
      dribblesAttempts,
      dribblesSuccess,
      foulsDrawn,
      foulsCommited,
      cardsYellow,
      cardsYellowRed,
      cardsRed,
      penaltyScored,
      penaltyMissed,
      penaltySaved
    ];
    sheet.getRange(row, 1, 1, data.length).setValues([data]);

    row += 1;
  }
}

function insertImage(): void {
  // シートを取得
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  for (let i = 765; i <= lastRow; i++) {
    const photo = sheet.getRange(i, 1).getValue();
    const photoData = SpreadsheetApp.newCellImage().setSourceUrl(photo).build();
    const teamLogo = sheet.getRange(i, 10).getValue();
    const teamLogoData = SpreadsheetApp.newCellImage().setSourceUrl(teamLogo).build();
    sheet.getRange(i, 1).setValue(photoData);
    sheet.getRange(i, 10).setValue(teamLogoData);
  }
}