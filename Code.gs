const API_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiIsImtpZCI6IjI4YTMxOGY3LTAwMDAtYTFlYi03ZmExLTJjNzQzM2M2Y2NhNSJ9.eyJpc3MiOiJzdXBlcmNlbGwiLCJhdWQiOiJzdXBlcmNlbGw6Z2FtZWFwaSIsImp0aSI6ImY2OWYyMWI0LTI0ZDctNDU3YS1iNmVlLWQwZmYyMzFiODBmNCIsImlhdCI6MTc3MzU0MTMzOCwic3ViIjoiZGV2ZWxvcGVyLzMwZWUzNGRjLTQ3MWItYTI0Mi0yMzdkLTQxZjQ4M2YwY2I3YSIsInNjb3BlcyI6WyJjbGFzaCJdLCJsaW1pdHMiOlt7InRpZXIiOiJkZXZlbG9wZXIvc2lsdmVyIiwidHlwZSI6InRocm90dGxpbmcifSx7ImNpZHJzIjpbIjQ1Ljc5LjIxOC43OSJdLCJ0eXBlIjoiY2xpZW50In1dfQ.xS-vUuYz1uPrfl0Funi8FUTID9Nvl_fcpp6lALVjgtUKQTjgI6Z-ydZQ7MEL_Xne5YS-xvSW0X24nVenZvluww";
const CLAN_TAG = "%23YQP0GJ9";
const BASE_URL = "https://cocproxy.royaleapi.dev/v1";

// NEW: The API Router. This replaces the old HTML service.
function doGet(e) {
  // If no action is provided, return a friendly error
  if (!e || !e.parameter || !e.parameter.action) {
    return ContentService.createTextOutput(JSON.stringify({ error: "No action specified." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = e.parameter.action;
  let responseData = {};

  try {
    if (action === 'getDashboardData') {
      responseData = getDashboardData();
    } else if (action === 'syncClanData') {
      syncClanData();
      responseData = { status: "success", message: "Clan data synced successfully." };
    } else {
      responseData = { error: "Unknown action requested." };
    }
  } catch (error) {
    responseData = { error: error.message };
  }

  // Return the data as a stringified JSON object
  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

function cleanBuggedData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const compIdx = data[0].indexOf("Completion %");
  const heroIdx = data[0].indexOf("Hero Total");
  if (compIdx === -1 || heroIdx === -1) return;

  let deletedCount = 0;
  for (let i = data.length - 1; i > 0; i--) {
    let pct = parseFloat(data[i][compIdx]);
    let heroStr = String(data[i][heroIdx]);
    if (pct > 100 || heroStr.includes("/305")) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  console.log(`Successfully cleaned ${deletedCount} corrupted rows.`);
}

function setupMollyMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Data")) ss.getActiveSheet().setName("Data");
  if (!ss.getSheetByName("Activity_Logs")) ss.insertSheet("Activity_Logs");
  if (!ss.getSheetByName("Capital_Raids")) ss.insertSheet("Capital_Raids");
  if (!ss.getSheetByName("War_Data")) ss.insertSheet("War_Data");
  if (!ss.getSheetByName("CWL_History")) ss.insertSheet("CWL_History");
  if (!ss.getSheetByName("Activity_Timeline")) ss.insertSheet("Activity_Timeline");
}

function syncClanData() {
  if (API_TOKEN === "PASTE_YOUR_API_TOKEN_HERE") throw new Error("Please paste your API Token.");

  const options = { "method": "GET", "headers": { "Authorization": "Bearer " + API_TOKEN }, "muteHttpExceptions": true };

  const membersRes = UrlFetchApp.fetch(`${BASE_URL}/clans/${CLAN_TAG}`, options);
  if (membersRes.getResponseCode() !== 200) return;
  const membersData = JSON.parse(membersRes.getContentText()).memberList;

  fetchCapitalRaids(options);
  fetchWarData(options);
  updateCwlDatabase(options);

  const requests = membersData.map(m => ({
    url: `${BASE_URL}/players/${m.tag.replace("#", "%23")}`,
    headers: { "Authorization": "Bearer " + API_TOKEN },
    muteHttpExceptions: true
  }));

  const responses = UrlFetchApp.fetchAll(requests);
  const now = new Date();

  responses.forEach(res => {
    if (res.getResponseCode() === 200) {
      const playerData = JSON.parse(res.getContentText());
      processOfficialApiData(playerData, now);
      updateActivityLog(playerData, now);
    }
  });

  CacheService.getScriptCache().remove("mollyDashboardData");
}

function processOfficialApiData(data, now) {
  const tag = data.tag;
  let extractedLevels = {};
  let currentTotal = 0, heroTotal = 0, petTotal = 0, labTotal = 0;

  const heroMap = { "Barbarian King": 105, "Archer Queen": 105, "Grand Warden": 80, "Royal Champion": 55, "Minion Prince": 95, "Dragon Duke": 25 };
  if (data.heroes) {
    data.heroes.forEach(hero => {
      if (hero.village === "home" && heroMap[hero.name]) { currentTotal += hero.level; heroTotal += hero.level; extractedLevels[hero.name] = hero.level; }
    });
  }

  const petMap = { "Electro Owl": 15, "Unicorn": 15, "Frosty": 15, "Diggy": 10, "Phoenix": 10, "Spirit Fox": 10, "Angry Jelly": 10, "Sneezy": 10, "Greedy Raven": 10 };
  if (data.troops) {
    data.troops.forEach(t => {
      if (t.village === "home") {
        if (petMap[t.name]) { currentTotal += t.level; petTotal += t.level; extractedLevels[t.name] = t.level; } else { labTotal += t.level; }
      }
    });
  }
  if (data.spells) data.spells.forEach(s => { if(s.village === "home") labTotal += s.level; });

  let equipmentList = [];
  if (data.heroEquipment) {
    equipmentList = data.heroEquipment.map(eq => ({ name: eq.name, level: eq.level }));
  }

  extractedLevels["Timestamp"] = now.getTime();
  extractedLevels["Username"] = data.name;
  extractedLevels["TH"] = data.townHallLevel || 0; 
  extractedLevels["Lab Total"] = labTotal;
  extractedLevels["Completion %"] = ((currentTotal / 570) * 100).toFixed(2);
  extractedLevels["Hero Total"] = `${heroTotal}/465`;
  extractedLevels["Pet Total"] = `${petTotal}/105`;
  extractedLevels["Donations"] = data.donations || 0;
  extractedLevels["Donations Received"] = data.donationsReceived || 0;
  extractedLevels["War Stars"] = data.warStars || 0;
  extractedLevels["Equipment"] = JSON.stringify(equipmentList);

  updateDataSheet(tag, extractedLevels);
}

function updateDataSheet(tag, extractedLevels) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  const requiredHeaders = ["Tag", "Username", "Timestamp", "TH", "Completion %", "Hero Total", "Pet Total", "Lab Total", "Barbarian King", "Archer Queen", "Grand Warden", "Royal Champion", "Minion Prince", "Dragon Duke", "Electro Owl", "Unicorn", "Frosty", "Diggy", "Phoenix", "Spirit Fox", "Angry Jelly", "Sneezy", "Greedy Raven", "Donations", "Donations Received", "War Stars", "Equipment"];
  if (!headers[0]) headers = [];
  requiredHeaders.forEach(req => { if (headers.indexOf(req) === -1) { headers.push(req); sheet.getRange(1, headers.length).setValue(req); } });

  let targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 1).setValue(tag);
  for (let colIndex = 1; colIndex < headers.length; colIndex++) {
    if (extractedLevels[headers[colIndex]] !== undefined) sheet.getRange(targetRow, colIndex + 1).setValue(extractedLevels[headers[colIndex]]);
  }
}

function updateActivityLog(playerData, now) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activity_Logs");
  const timelineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activity_Timeline");
  const tag = playerData.tag; const dataRange = sheet.getDataRange().getValues(); let rowIndex = -1;
  for (let i = 1; i < dataRange.length; i++) { if (dataRange[i][0] === tag) { rowIndex = i + 1; break; } }
  const currentStats = { donations: playerData.donations || 0, received: playerData.donationsReceived || 0, attacks: playerData.attackWins || 0 };

  if (rowIndex === -1) {
    sheet.appendRow([tag, playerData.name, now.getTime(), currentStats.donations, currentStats.received, currentStats.attacks, 1]);
    if(timelineSheet) timelineSheet.appendRow([now.getTime()]);
  } else {
    const prev = { lastSeen: dataRange[rowIndex-1][2], donations: dataRange[rowIndex-1][3], received: dataRange[rowIndex-1][4], attacks: dataRange[rowIndex-1][5], score: dataRange[rowIndex-1][6] || 0 };
    let hasChanged = (currentStats.donations !== prev.donations) || (currentStats.received !== prev.received) || (currentStats.attacks !== prev.attacks);
    let newLastSeen = hasChanged ? now.getTime() : prev.lastSeen; let newScore = hasChanged ? prev.score + 1 : prev.score;
    sheet.getRange(rowIndex, 1, 1, 7).setValues([[tag, playerData.name, newLastSeen, currentStats.donations, currentStats.received, currentStats.attacks, newScore]]);

    if (hasChanged && timelineSheet) timelineSheet.appendRow([now.getTime()]);
  }
}

function fetchCapitalRaids(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Capital_Raids");
  const res = UrlFetchApp.fetch(`${BASE_URL}/clans/${CLAN_TAG}/capitalraidseasons`, options);
  if (res.getResponseCode() !== 200) return;
  const data = JSON.parse(res.getContentText()); if(!data.items || data.items.length === 0) return;

  sheet.clear(); sheet.appendRow(["Tag", "Name", "Attacks", "AttackLimit", "BonusAttacks", "Looted", "HISTORY_JSON"]);
  let historyData = data.items.map(season => {
    let d = season.startTime.substring(0,8); return { date: d.substring(4,6) + "/" + d.substring(6,8) + "/" + d.substring(2,4), loot: season.capitalTotalLoot };
  }).reverse();
  sheet.getRange(1, 7).setValue(JSON.stringify(historyData));

  (data.items[0].members || []).forEach(m => { sheet.appendRow([m.tag, m.name, m.attacks, m.attackLimit, m.bonusAttackLimit, m.capitalResourcesLooted]); });
}

function updateCwlDatabase(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CWL_History");
  if (!sheet) return;
  let res = UrlFetchApp.fetch(`${BASE_URL}/clans/${CLAN_TAG}/currentwar/leaguegroup`, options);
  if (res.getResponseCode() === 200) {
    let group = JSON.parse(res.getContentText());
    if (group.state !== "notInWar" && group.season) {
      let dataRange = sheet.getDataRange().getValues(); let rowIndex = -1;
      for (let i = 0; i < dataRange.length; i++) { if (dataRange[i][0] === group.season) { rowIndex = i + 1; break; } }
      let payload = JSON.stringify({ season: group.season, state: group.state });
      if (rowIndex === -1) { sheet.appendRow([group.season, payload]); } else { sheet.getRange(rowIndex, 2).setValue(payload); }
    }
  }
}

// 👑 THE ESPIONAGE ENGINE: Deep CWL Analytics & Target Patterns
function fetchWarData(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("War_Data"); sheet.clear();
  let currentWarInfo = { type: "NONE", data: null, cwlDetails: [], myCwlStats: {} };
  let res = UrlFetchApp.fetch(`${BASE_URL}/clans/${CLAN_TAG}/currentwar/leaguegroup`, options);
  if (res.getResponseCode() === 200) {
    let group = JSON.parse(res.getContentText());
    if (group.state !== "notInWar") {
      currentWarInfo.type = "CWL";
      currentWarInfo.data = { season: group.season, state: group.state };

      let cwlWars = [];
      let myClanTag = decodeURIComponent(CLAN_TAG);
      let enemyAttackPatterns = {};
      let myCwlStats = {};

      if (group.rounds) {
        for(let round of group.rounds) {
          if(round.warTags[0] === "#0") continue;
          for(let wTag of round.warTags) {
            if(wTag === "#0") continue;
            let wRes = UrlFetchApp.fetch(`${BASE_URL}/clanwarleagues/wars/${wTag.replace("#", "%23")}`, options);
            if (wRes.getResponseCode() === 200) {
              let wData = JSON.parse(wRes.getContentText());

              // ESPIONAGE: Analyze attack patterns across ALL wars to build a profile for every clan
              [wData.clan, wData.opponent].forEach(clan => {
                if (!enemyAttackPatterns[clan.tag]) {
                  enemyAttackPatterns[clan.tag] = { mirrors: 0, dips: 0, reaches: 0, total: 0, threeStars: 0 };
                }
                let p = enemyAttackPatterns[clan.tag];
                let oppClan = clan.tag === wData.clan.tag ? wData.opponent : wData.clan;

                (clan.members || []).forEach(m => {
                  if (m.attacks) {
                    m.attacks.forEach(atk => {
                      let defender = (oppClan.members || []).find(d => d.tag === atk.defenderTag);
                      if (defender) {
                        p.total++;
                        if (atk.stars === 3) p.threeStars++;

                        // Position 1 is strongest, 15 is weakest.
                        if (m.mapPosition === defender.mapPosition) p.mirrors++;
                        else if (m.mapPosition < defender.mapPosition) p.dips++;
                        else p.reaches++;
                      }
                    });
                  }
                });
              });

              // ANALYTICS: Process our specific war in this round
              if (wData.clan.tag === myClanTag || wData.opponent.tag === myClanTag) {
                let enemyClan = wData.clan.tag === myClanTag ? wData.opponent : wData.clan;
                let myClan = wData.clan.tag === myClanTag ? wData.clan : wData.opponent;

                // Calculate True Hit Rates & Defensive Liabilities for our roster
                (myClan.members || []).forEach(m => {
                  if (!myCwlStats[m.tag]) myCwlStats[m.tag] = { attacks: 0, threeStars: 0, defenses: 0, defStars: 0 };
                  if (m.attacks) {
                    m.attacks.forEach(atk => {
                      myCwlStats[m.tag].attacks++;
                      if (atk.stars === 3) myCwlStats[m.tag].threeStars++;
                    });
                  }
                });

                // Look at the enemy attacks to see who on our team was hit
                (enemyClan.members || []).forEach(m => {
                  if (m.attacks) {
                    m.attacks.forEach(atk => {
                      if (!myCwlStats[atk.defenderTag]) myCwlStats[atk.defenderTag] = { attacks: 0, threeStars: 0, defenses: 0, defStars: 0 };
                      myCwlStats[atk.defenderTag].defenses++;
                      myCwlStats[atk.defenderTag].defStars += atk.stars;
                    });
                  }
                });

                // Condense Map Roster arrays to save JSON space
                let enemyLineup = (enemyClan.members || []).map(m => ({
                  name: m.name, tag: m.tag, th: m.townhallLevel, mapPosition: m.mapPosition, attacks: m.attacks ? m.attacks.length : 0
                })).sort((a,b) => a.mapPosition - b.mapPosition);

                let myLineup = (myClan.members || []).map(m => ({
                  name: m.name, tag: m.tag, th: m.townhallLevel, mapPosition: m.mapPosition
                })).sort((a,b) => a.mapPosition - b.mapPosition);

                cwlWars.push({
                  state: wData.state,
                  opponent: { name: enemyClan.name, tag: enemyClan.tag },
                  enemyLineup: enemyLineup,
                  myLineup: myLineup
                });
              }
            }
          }
        }
      }

      // Assign the final computed attack patterns to each opponent in the lineup
      cwlWars.forEach(w => {
        w.attackPattern = enemyAttackPatterns[w.opponent.tag] || { mirrors: 0, dips: 0, reaches: 0, total: 0, threeStars: 0 };
      });

      currentWarInfo.cwlDetails = cwlWars;
      currentWarInfo.myCwlStats = myCwlStats;
    }
  }
  if (currentWarInfo.type === "NONE") {
    res = UrlFetchApp.fetch(`${BASE_URL}/clans/${CLAN_TAG}/currentwar`, options);
    if (res.getResponseCode() === 200) {
      let war = JSON.parse(res.getContentText());
      if (war.state !== "notInWar") {
        currentWarInfo.type = "REGULAR";
        currentWarInfo.data = { state: war.state, opponent: { name: war.opponent ? war.opponent.name : "Unknown" } };
      }
    }
  }
  
  let jsonPayload = JSON.stringify(currentWarInfo);
  if (jsonPayload.length > 49000) {
    // Ultimate Failsafe - Trims older rounds if payload breaks 50K
    currentWarInfo.cwlDetails = currentWarInfo.cwlDetails.slice(-3);
    jsonPayload = JSON.stringify(currentWarInfo);
  }
  sheet.getRange(1, 1).setValue(jsonPayload);
}

function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("mollyDashboardData");
  if (cachedData) return JSON.parse(cachedData);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data").getDataRange().getValues();
  if (!dataSheet || dataSheet.length === 0 || !dataSheet[0] || dataSheet[0].length === 0) {
    return { players: [], clanPct: 0, clanHeroes: "0/0", clanPets: "0/0", clanDonations: 0, clanReceived: 0, clanHistoryDates: [], clanHistoryPcts: [], clanCapitalLoot: 0, capitalHistory: [], cwlHistory: [], warInfo: { type: "NONE", data: null, cwlDetails: [], myCwlStats: {} }, activityGraphs: { hourly: [], daily: [] } };
  }

  const activityData = ss.getSheetByName("Activity_Logs").getDataRange().getValues();
  const capitalData = ss.getSheetByName("Capital_Raids").getDataRange().getValues();
  let cwlHistoryArray = [];
  try {
    let cwlDbData = ss.getSheetByName("CWL_History").getDataRange().getValues();
    for (let i = 0; i < cwlDbData.length; i++) { if (cwlDbData[i][1]) cwlHistoryArray.push(JSON.parse(cwlDbData[i][1])); }
  } catch(e) {}
  cwlHistoryArray.reverse();

  let hourlyActivity = new Array(24).fill(0);
  let dailyActivity = new Array(7).fill(0);
  try {
    let timelineData = ss.getSheetByName("Activity_Timeline").getDataRange().getValues();
    const tz = "America/Chicago";
    for (let i = 1; i < timelineData.length; i++) {
      let ts = timelineData[i][0];
      if (!ts) continue;
      let d = new Date(ts);
      let hour = parseInt(Utilities.formatDate(d, tz, "H"));
      let dayRaw = parseInt(Utilities.formatDate(d, tz, "u"));
      let dayIdx = dayRaw === 7 ? 0 : dayRaw;
      hourlyActivity[hour]++;
      dailyActivity[dayIdx]++;
    }
  } catch(e) {}

  let warRaw = ""; try { warRaw = ss.getSheetByName("War_Data").getRange(1, 1).getValue(); } catch(e) {}
  let warParsed = { type: "NONE", data: null, cwlDetails: [], myCwlStats: {} };
  try { if (warRaw) warParsed = JSON.parse(warRaw); } catch(e) {}

  let activityMap = {}; if (activityData && activityData.length > 1) { for(let i=1; i<activityData.length; i++) activityMap[activityData[i][0]] = { lastSeenTs: activityData[i][2], activityScore: activityData[i][6] }; }
  
  let capitalMap = {}; let totalClanLoot = 0; if (capitalData && capitalData.length > 1) { for(let i=1; i<capitalData.length; i++) { capitalMap[capitalData[i][0]] = { attacks: capitalData[i][2], limit: capitalData[i][3] + capitalData[i][4], looted: capitalData[i][5] };
  totalClanLoot += capitalData[i][5]; } }
  
  let capitalHistory = []; try { let histStr = ss.getSheetByName("Capital_Raids").getRange(1, 7).getValue(); if (histStr) capitalHistory = JSON.parse(histStr); } catch(e) {}

  let playerMap = {};
  const headers = dataSheet[0];
  const tagIdx = 0; const nameIdx = headers.indexOf("Username"); const tsIdx = headers.indexOf("Timestamp"); const compIdx = headers.indexOf("Completion %");

  for (let i = 1; i < dataSheet.length; i++) {
    const row = dataSheet[i]; const tag = row[tagIdx]; if (!tag || tag === "Tag") continue;
    const ts = tsIdx > -1 && row[tsIdx] ? row[tsIdx] : 0;
    const dateLabel = ts > 0 ? Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), "MM/dd/yy") : "";
    if (!playerMap[tag]) playerMap[tag] = { tag: tag, history: [], latestTs: 0 };
    playerMap[tag].history.push({ date: dateLabel, pct: parseFloat(row[compIdx]) || 0, ts: ts, rawRow: row });
    if (ts >= playerMap[tag].latestTs || playerMap[tag].latestTs === 0) playerMap[tag].latestTs = ts;
  }

  let results = []; let uniqueDatesMap = {};
  const thirtyDaysAgo = new Date().getTime() - (30 * 24 * 60 * 60 * 1000);
  const heroNames = ["Barbarian King", "Archer Queen", "Grand Warden", "Royal Champion", "Minion Prince", "Dragon Duke"];
  const petNames = ["Electro Owl", "Unicorn", "Frosty", "Diggy", "Phoenix", "Spirit Fox", "Angry Jelly", "Sneezy", "Greedy Raven"];

  let clanTotalDonations = 0;
  let clanTotalReceived = 0;

  for (let key in playerMap) {
    let p = playerMap[key]; p.history.sort((a, b) => a.ts - b.ts);
    let row = p.history[p.history.length - 1].rawRow; let prevRow = p.history.length > 1 ? p.history[p.history.length - 2].rawRow : row;
    p.history.forEach(h => { if (!uniqueDatesMap[h.date] || h.ts > uniqueDatesMap[h.date]) uniqueDatesMap[h.date] = h.ts; });
    let chartHistory = p.history.filter(h => h.ts >= thirtyDaysAgo); if (chartHistory.length === 0 && p.history.length > 0) chartHistory.push(p.history[p.history.length - 1]);
    let act = activityMap[key] || { lastSeenTs: 0, activityScore: 0 }; let cap = capitalMap[key] || { attacks: 0, limit: 6, looted: 0 };

    let donations = headers.indexOf("Donations") > -1 ? parseInt(row[headers.indexOf("Donations")]) || 0 : 0;
    let received = headers.indexOf("Donations Received") > -1 ? parseInt(row[headers.indexOf("Donations Received")]) || 0 : 0;
    clanTotalDonations += donations;
    clanTotalReceived += received;

    let cwlStats = warParsed.myCwlStats[p.tag] || { attacks: 0, threeStars: 0, defenses: 0, defStars: 0 };

    let playerObj = {
      name: nameIdx > -1 && row[nameIdx] ? row[nameIdx] : p.tag, tag: p.tag,
      thLevel: headers.indexOf("TH") > -1 ? parseInt(row[headers.indexOf("TH")]) : 0, 
      percentage: parseFloat((row[compIdx] || 0).toString()).toFixed(2), prevPct: parseFloat((prevRow[compIdx] || 0).toString()).toFixed(2),
      heroStats: headers.indexOf("Hero Total") > -1 ? row[headers.indexOf("Hero Total")] : "0/465", petStats: headers.indexOf("Pet Total") > -1 ? row[headers.indexOf("Pet Total")] : "0/105",
      labStats: headers.indexOf("Lab Total") > -1 ? row[headers.indexOf("Lab Total")] : 0,
      donations: donations, donationsReceived: received, warStars: headers.indexOf("War Stars") > -1 ? row[headers.indexOf("War Stars")] : 0,
      lastSeenTs: act.lastSeenTs, activityScore: act.activityScore, capitalAttacks: `${cap.attacks}/${cap.limit}`, capitalLooted: cap.looted,
      cwlStats: cwlStats, 
      historyDates: chartHistory.map(h => h.date), historyPcts: chartHistory.map(h => h.pct),
      heroes: {}, pets: {}, upgrades: [], recentlyCompleted: []
    };

    const getLvl = (r, name) => { let idx = headers.indexOf(name); return idx > -1 ? parseInt(r[idx]) || 0 : 0; };
    
    heroNames.forEach(h => { playerObj.heroes[h] = getLvl(row, h); if(playerObj.heroes[h] > getLvl(prevRow, h)) playerObj.recentlyCompleted.push(`${h} to Lvl ${playerObj.heroes[h]}`); });
    petNames.forEach(pet => { playerObj.pets[pet] = getLvl(row, pet); if(playerObj.pets[pet] > getLvl(prevRow, pet)) playerObj.recentlyCompleted.push(`${pet} to Lvl ${playerObj.pets[pet]}`); });
    results.push(playerObj);
  }

  results.sort((a, b) => b.percentage - a.percentage);
  let clanTotalHeroes = 0, clanMaxHeroes = 0; let clanTotalPets = 0, clanMaxPets = 0;
  results.forEach(p => {
    clanTotalHeroes += parseInt((p.heroStats || "0/465").split('/')[0]) || 0; clanMaxHeroes += parseInt((p.heroStats || "0/465").split('/')[1]) || 465;
    clanTotalPets += parseInt((p.petStats || "0/105").split('/')[0]) || 0; clanMaxPets += parseInt((p.petStats || "0/105").split('/')[1]) || 105;
  });

  let uniqueDatesArray = Object.keys(uniqueDatesMap).map(date => ({ label: date, ts: uniqueDatesMap[date] })).sort((a, b) => a.ts - b.ts);
  let clanHistoryDates = [], clanHistoryPcts = [];
  uniqueDatesArray.filter(d => d.ts >= thirtyDaysAgo).forEach(d => {
    let totalPctAtDate = 0, activePlayerCount = 0;
    for (let key in playerMap) {
      let validRecords = playerMap[key].history.filter(h => h.ts <= (d.ts + 86400000));
      if (validRecords.length > 0) { totalPctAtDate += validRecords[validRecords.length - 1].pct; activePlayerCount++; }
    }
    if (activePlayerCount > 0) { clanHistoryDates.push(d.label); clanHistoryPcts.push(parseFloat((totalPctAtDate / activePlayerCount).toFixed(2))); }
  });

  const finalResponse = {
    players: results, clanPct: clanHistoryPcts.length > 0 ? clanHistoryPcts[clanHistoryPcts.length - 1] : 0,
    clanHeroes: `${clanTotalHeroes}/${clanMaxHeroes}`, clanPets: `${clanTotalPets}/${clanMaxPets}`,
    clanDonations: clanTotalDonations, clanReceived: clanTotalReceived,
    clanHistoryDates: clanHistoryDates, clanHistoryPcts: clanHistoryPcts,
    clanCapitalLoot: totalClanLoot, capitalHistory: capitalHistory,
    cwlHistory: cwlHistoryArray,
    warInfo: { type: warParsed.type, data: warParsed.data, cwlDetails: warParsed.cwlDetails },
    activityGraphs: { hourly: hourlyActivity, daily: dailyActivity }
  };
  try { CacheService.getScriptCache().put("mollyDashboardData", JSON.stringify(finalResponse), 1800); } catch(e) { console.warn("Cache bypassed due to size."); }
  return finalResponse;
}
