// ================= CONFIG =================

const BASE_URL = "https://fantasy.iplt20.com";

const BOOSTER_NAME_MAP = {
  1: "Wild Card",
  3: "Double Power",
  9: "Foreign Stars",
  10: "Indian Warriors",
  11: "Free Hit",
  12: "Triple Captain"
};

const BOOSTER_AVG = {
  DOUBLE_POWER: 1100,
  FOREIGN_STARS: 600,
  INDIAN_WARRIORS: 800,
  TRIPLE_CAPTAIN: 180
};

const TEAM_COLORS = {};
const COLOR_POOL = [
  "#1f77b4","#d62728","#2ca02c","#9467bd",
  "#ff7f0e","#8c564b","#e377c2","#7f7f7f",
  "#bcbd22","#17becf","#393b79","#637939"
];

// ================= UI HELPERS =================

function getTeamColor(teamName){
  if (TEAM_COLORS[teamName]) return TEAM_COLORS[teamName];
  const i = Object.keys(TEAM_COLORS).length % COLOR_POOL.length;
  return TEAM_COLORS[teamName] = COLOR_POOL[i];
}

function applyTeamStyle(sheet,row,col,team){
  sheet.getRange(row,col)
    .setFontColor(getTeamColor(team))
    .setFontWeight("bold");
}

function resetSheet(sheet){
  const r = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns());

  r.clearContent();
  r.setBorder(false,false,false,false,false,false);
  r.setBackground(null);
  r.setFontWeight("normal");
  r.setFontColor("#000000");
  r.setHorizontalAlignment("left");
}

// ================= MAIN =================

function runTracker(){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  resetSheet(sheet);

  const ctx = fetchData();

  const maps = buildMaps(ctx);

  const boosterStatsMap = buildBoosterStatsFromTable1(ctx, maps);

  const t1 = renderTable1(sheet,ctx,maps);
  const t2 = renderTable2(sheet,ctx,maps,t1);
  const t3 = renderTable3(sheet,ctx,maps,t2, boosterStatsMap);
  const t4 = renderTable4(sheet,ctx,maps,t3);
  renderTable5(sheet,ctx,maps,t4);
}

// ================= DATA =================

function fetchData(){

  const cookie = "my11c-uid=4065560865; my11c-authToken=eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJhdXRoIiwicHJvZHVjdF90eXBlIjoyLCJzZXNzaW9uSWQiOiI2dEgySUkyZUtyNGVMeHhWL1RSMTd4MVg2ODFSSUJzWjFRUHVJUm1sdEFqVEh0a2dxbzd6MlRjN1A3elFOT3pXIiwidXNlcklkIjoxNzY3MzAxNDMsIndoYXRzYXBwQ2FsbCI6ZmFsc2UsImlhdCI6MTc3NTc2MjExMiwiZXhwIjoxNzc1NzY5MzEyfQ.HBlynyDxnOTUJpIQg8423khXJXWNPQGfwnAg2xHiKGI; my11_classic_game=%7B%0A%20%20%22UserName%22%3A%20%22Dada%20Dragons%22%2C%0A%20%20%22HasTeam%22%3A%201%2C%0A%20%20%22TeamName%22%3A%20%22Dada%20Dragons%22%2C%0A%20%20%22FavTeamId%22%3A%20%221106%22%2C%0A%20%20%22SocialId%22%3A%20%224065560865%22%2C%0A%20%20%22GUID%22%3A%20%224ee8faf4-2485-11f1-b171-06a69fcb782b%22%2C%0A%20%20%22ActiveTour%22%3A%20null%2C%0A%20%20%22IsTourActive%22%3A%200%2C%0A%20%20%22UserId%22%3A%20%22E924AB7811605C8B%22%2C%0A%20%20%22TeamId%22%3A%20%22E924AB7811605C8B%22%2C%0A%20%20%22ProfileURL%22%3A%20%22%22%2C%0A%20%20%22TeamName_Allow%22%3A%20%220%22%2C%0A%20%20%22Version%22%3A%20%224%22%2C%0A%20%20%22IsIndian%22%3A%20%221%22%0A%7D; _gid=GA1.2.1373948057.1775762113; TEAM_BUSTER=20260409195458; PUBLIC_LEAGUE_BUSTER=20260409195458; PRIVATE_LEAGUE_BUSTER=20260409195458; badge=20260409195458; profile=20260409195458; _ga=GA1.1.2059026392.1775762113; _ga_H403WMH7VL=GS2.1.s1775766872$o2$g1$t1775767638$j60$l0$h0; _ga_N6TKWNZRZ6=GS2.1.s1775766872$o2$g1$t1775767638$j60$l0$h0";

  const fixtures = fetchFixtures(cookie);
  const gamedayId = getCurrentGameDayId(fixtures);

  const leaderboard = fetchLeaderboard(cookie,gamedayId);
  const teams = (leaderboard?.Data?.Value || []).filter(t=>t && t.islocked!==2);

  const allGameDays = getAllGameDayIds(fixtures);

  return {
    cookie,
    gamedayId,
    teams,
    allGameDays,
    playerMapsByGameDay: buildPlayerMapsByGameDay(cookie,allGameDays),
    teamCache: buildTeamDataCache(cookie,teams,allGameDays),
    playerMap: buildPlayerMap(fetchPlayers(cookie,gamedayId))
  };
}

// ================= MAPS =================

function buildMaps(ctx){

  const {teams,allGameDays,playerMapsByGameDay,teamCache} = ctx;

  const matchData = buildMatchWiseData(teams,allGameDays,playerMapsByGameDay,teamCache);

  const table2Map = {};
  matchData.forEach(r=>{
    const team=r[0];
    const map={};
    let m=1;
    for(let i=1;i<r.length;i+=2){
      map[m]={captain:r[i]||0,vc:r[i+1]||0};
      m++;
    }
    table2Map[team]=map;
  });

  const totalData = buildMatchWiseTotalAndSubs(
    teams,allGameDays,playerMapsByGameDay,teamCache
  );

  const table4Map = {};
  totalData.forEach(r=>table4Map[r[0]]=r);

  return {matchData,totalData,table2Map,table4Map};
}

// ================= TABLE 2 =================

function renderTable2(sheet,ctx,maps,layout){

  const startRow = layout.end + 3;
  const header = buildMatchHeader(ctx.allGameDays);

  sheet.getRange(startRow,1,1,header.length)
     .setFontWeight("bold")
     .setBackground("#f1f3f4");

  sheet.getRange(startRow,1,1,header.length).setValues([header]);
  sheet.getRange(startRow+1,1,maps.matchData.length,header.length)
    .setValues(maps.matchData);

  maps.matchData.forEach((r,i)=>{
    applyTeamStyle(sheet,startRow+1+i,1,r[0]);
  });

  sheet.getRange(startRow,1,maps.matchData.length+1,header.length)
     .setBorder(true,true,true,true,true,true);

  return { end:startRow + maps.matchData.length };
}

// ================= TABLE 3 =================

function renderTable3(sheet,ctx,maps,layout,boosterStatsMap){

  const startRow = layout.end + 2;

  const playerNameMap = buildPlayerNameMap(ctx.playerMapsByGameDay);
  const teamHistory = buildTeamHistoryFromCache(
    ctx.teamCache,ctx.teams,ctx.allGameDays
  );

  const boosterHeaders = [
    "Best Wild Card",
    "Best Double Power",
    "Best Foreign Stars",
    "Best Indian Warriors",
    "Best Free Hit",
    "Best Triple Captain"
  ];

  const header = [
    "Team","Avg C","Avg VC","Avg (C+VC)","Best Captain","Best VC",
    ...boosterHeaders
  ];

  const baseData = buildAverageTable(
    maps.matchData,
    teamHistory,
    playerNameMap
  );

  const data = baseData.map(row => {

    const teamName = row[0];

    const boosterStats = boosterStatsMap[teamName] || {};

    return [
      ...row,
      boosterStats["Wild Card"] || "",
      boosterStats["Double Power"] || "",
      boosterStats["Foreign Stars"] || "",
      boosterStats["Indian Warriors"] || "",
      boosterStats["Free Hit"] || "",
      boosterStats["Triple Captain"] || ""
    ];
  });
  
  sheet.getRange(startRow,1,1,header.length)
      .setFontWeight("bold")
      .setBackground("#f1f3f4");
  sheet.getRange(startRow,1,1,header.length).setValues([header]);
  sheet.getRange(startRow+1,1,data.length,header.length).setValues(data);

  data.forEach((r,i)=>{
    if(r[0]) applyTeamStyle(sheet,startRow+1+i,1,r[0]);
  });

  sheet.getRange(startRow,1,data.length+1,header.length)
     .setBorder(true,true,true,true,true,true);

  return { end:startRow + data.length };
}

// ================= TABLE 4 =================

function renderTable4(sheet,ctx,maps,layout){

  const startRow = layout.end + 2;

  let header=["Team Name"];
  ctx.allGameDays.forEach(g=>{
    header.push(`G${g} Total`,`G${g} Subs`,`G${g} Players`,`G${g} Pts/Player`);
  });

  sheet.getRange(startRow,1,1,header.length)
     .setFontWeight("bold")
     .setBackground("#f1f3f4");

  sheet.getRange(startRow,1,1,header.length).setValues([header]);
  sheet.getRange(startRow+1,1,maps.totalData.length,header.length)
    .setValues(maps.totalData);

  maps.totalData.forEach((r,i)=>{
    applyTeamStyle(sheet,startRow+1+i,1,r[0]);
  });

  sheet.getRange(startRow,1,maps.totalData.length+1,header.length)
     .setBorder(true,true,true,true,true,true);

  return { end:startRow + maps.totalData.length };
}

// ================= TABLE 1 =================

function renderTable1(sheet, ctx, maps) {

  const header = [
    "Rank","Team Name","Points","Sub Used","Sub Left","Credits Left",
    "Subs Efficiency","Avg Pts/Player/Match","Projected Points","Projected Rank","Subs Plan",
    "Captain Name","Captain Score","VC Name","VC Score","Booster"
  ];

  sheet.getRange(1,1,1,header.length).setValues([header]);

  const rows = [];

  ctx.teams.forEach(t => {

    const val = ctx.teamCache[t.temname][ctx.gamedayId];
    if (!val) return;

    const captain = ctx.playerMap[val.mcapt] || {};
    const vc = ctx.playerMap[val.vcapt] || {};

    // ✅ Booster (enhanced)
    const booster = formatBoosterWithPoints(
      val,
      t.temname,
      maps.table4Map,
      maps.table2Map
    );

    // ✅ Subs Efficiency
    let subsEfficiency = 0;
    let projectedPoints = 0;

    if (val.subusr) {

      subsEfficiency = Number((t.points / val.subusr).toFixed(2));

      const base = (subsEfficiency * 160) + (subsEfficiency * 10);

      const usage = getBoosterUsageCount(val);
      const remaining = getRemainingBoosters(usage);

      const boosterPoints =
        remaining.DOUBLE_POWER * BOOSTER_AVG.DOUBLE_POWER +
        remaining.FOREIGN_STARS * BOOSTER_AVG.FOREIGN_STARS +
        remaining.INDIAN_WARRIORS * BOOSTER_AVG.INDIAN_WARRIORS +
        remaining.TRIPLE_CAPTAIN * BOOSTER_AVG.TRIPLE_CAPTAIN;

      projectedPoints = Number((base + boosterPoints).toFixed(2));
    }

    // ✅ Subs Plan
    const matchesPlayed = ctx.gamedayId;
    const matchesLeft = 70 - matchesPlayed;
    const subsPlan = buildBalancedSubsPlan(val.subleft, matchesLeft);

    // ✅ Avg Points Per Player Per Match
    let totalPtsPerPlayer = 0;
    let matchesCount = 0;

    ctx.allGameDays.forEach(gd => {

      const teamData = ctx.teamCache[t.temname];
      const valMatch = teamData ? teamData[gd] : null;

      if (!valMatch || !valMatch.plyid) return;

      const pmMatch = ctx.playerMapsByGameDay[gd] || {};

      let base = 0;
      let count = 0;

      valMatch.plyid.forEach(pid => {
        const p = pmMatch[pid];
        if (!p) return;

        base += Number(p.points || 0);
        count++;
      });

      const cM = pmMatch[valMatch.mcapt];
      const vcM = pmMatch[valMatch.vcapt];

      const total =
        base +
        (cM ? cM.points : 0) +
        (vcM ? vcM.points * 0.5 : 0);

      if (count > 0) {
        totalPtsPerPlayer += total / count;
        matchesCount++;
      }
    });

    const avgPtsPerPlayerMatch = matchesCount
      ? Number((totalPtsPerPlayer / matchesCount).toFixed(2))
      : 0;

    rows.push({
      row: [
        t.rank,
        t.temname,
        t.points,
        val.subusr,
        val.subleft,
        val.tembal,
        subsEfficiency,
        avgPtsPerPlayerMatch,
        projectedPoints,
        0, // placeholder for rank
        subsPlan,
        captain.name,
        (captain.points || 0) * 2,
        vc.name,
        (vc.points || 0) * 1.5,
        booster
      ],
      team: t.temname,
      projected: projectedPoints
    });
  });

  // ✅ Projected Rank
  const sorted = [...rows].sort((a,b)=>b.projected - a.projected);
  const rankMap = {};
  sorted.forEach((d,i)=> rankMap[d.team] = i+1);

  const finalData = rows.map(d=>{
    d.row[9] = rankMap[d.team];
    return d.row;
  });

  sheet.getRange(2,1,finalData.length,header.length).setValues(finalData);

  // 🎨 Styling
  finalData.forEach((r,i)=>{
    applyTeamStyle(sheet,2+i,2,r[1]);
  });

  sheet.getRange(1,1,1,header.length)
       .setBackground("#d9eaf7")
       .setFontWeight("bold");

  sheet.getRange(1,1,finalData.length+1,header.length)
     .setBorder(true,true,true,true,true,true);

  return { end: finalData.length + 1 };
}

// ================= TABLE 5 =================

function renderTable5(sheet,ctx,maps,layout){

  const startRow = layout.end + 2;

  const header=["Game","Best Team","Best Eff","Worst Team","Worst Eff"];
  sheet.getRange(startRow,1,1,header.length)
     .setFontWeight("bold")
     .setBackground("#f1f3f4")
     .setHorizontalAlignment("left");
  const data = buildEfficiencySummary(ctx.teams,ctx.allGameDays,maps.totalData);

  sheet.getRange(startRow,1,1,header.length).setValues([header]);
  sheet.getRange(startRow+1,1,data.length,header.length).setValues(data);

  data.forEach((r,i)=>{
    applyTeamStyle(sheet,startRow+1+i,2,r[1]);
    applyTeamStyle(sheet,startRow+1+i,4,r[3]);
  });

  sheet.getRange(startRow+1,1,data.length,header.length)
     .setHorizontalAlignment("left");
  sheet.getRange(startRow,1,data.length+1,header.length)
  .setBorder(true,true,true,true,true,true);
}

// ================= BOOSTER LOGIC =================

function formatBoosterWithPoints(val,team,table4Map,table2Map){

  if(!val.booster?.length) return "";

  const grouped={};

  val.booster.forEach(b=>{

    const name = BOOSTER_NAME_MAP[b.cf_boosterid] || "UNKNOWN";
    const g = b.cf_team_gamedayid;

    let pts=0;

    if(b.cf_boosterid===12){
      const cap = table2Map?.[team]?.[g]?.captain || 0;
      pts = (3*cap)/2;
    } else {
      const row = table4Map[team];
      if(row){
        const idx = 1 + (g-1)*4;
        pts = row[idx] || 0;
      }
    }

    if(!grouped[name]) grouped[name]=[];
    grouped[name].push(`G${g}: ${pts}`);
  });

  return Object.entries(grouped)
    .map(([k,v])=>`${k} (${v.join(", ")})`)
    .join(" | ");
}

//
// 🔧 HELPERS
//

function buildMatchWiseData(teams, gameDays, playerMaps, teamCache) {
  return teams.map(t => {
    let row = [t.temname];

    gameDays.forEach(gd => {
      const val = teamCache[t.temname][gd];
      if (!val) return row.push(0,0);

      const pm = playerMaps[gd] || {};
      const c = pm[val.mcapt];
      const vc = pm[val.vcapt];

      row.push(
        c ? (c.points||0)*2 : 0,
        vc ? (vc.points||0)*1.5 : 0
      );
    });

    return row;
  });
}

function buildTeamHistoryFromCache(teamCache, teams, gameDays) {
  const history = {};
  teams.forEach(t=>{
    history[t.temname] = gameDays.map(gd=>{
      const val = teamCache[t.temname][gd];
      return val ? {c:val.mcapt,vc:val.vcapt} : {};
    });
  });
  return history;
}

function getBoosterUsageCount(val){
  const u={DOUBLE_POWER:0,FOREIGN_STARS:0,INDIAN_WARRIORS:0,TRIPLE_CAPTAIN:0};
  val.booster?.forEach(b=>{
    if([1,3].includes(b.cf_boosterid))u.DOUBLE_POWER++;
    else if(b.cf_boosterid==9)u.FOREIGN_STARS++;
    else if(b.cf_boosterid==10)u.INDIAN_WARRIORS++;
    else if(b.cf_boosterid==12)u.TRIPLE_CAPTAIN++;
  });
  return u;
}

function getRemainingBoosters(u){
  const M=2;
  return {
    DOUBLE_POWER:Math.max(0,M-u.DOUBLE_POWER),
    FOREIGN_STARS:Math.max(0,M-u.FOREIGN_STARS),
    INDIAN_WARRIORS:Math.max(0,M-u.INDIAN_WARRIORS),
    TRIPLE_CAPTAIN:Math.max(0,M-u.TRIPLE_CAPTAIN)
  };
}

function buildPlayerMapsByGameDay(cookie, gameDays){
  const map={};
  gameDays.forEach(gd=>{
    try{
      map[gd]=buildPlayerMap(fetchPlayers(cookie,gd));
    }catch(e){ map[gd]={}; }
  });
  return map;
}

function buildTeamDataCache(cookie, teams, gameDays){
  const cache={};
  teams.forEach(t=>{
    cache[t.temname]={};
    gameDays.forEach(gd=>{
      try{
        const res=fetchTeam(cookie,t.temid,t.usrscoid,gd);
        cache[t.temname][gd]=res?.Data?.Value||null;
      }catch(e){ cache[t.temname][gd]=null; }
    });
  });
  return cache;
}

function buildPlayerMap(res) {
  const map = {};
  if (!res?.Data?.Value?.Players) return map;
  res.Data.Value.Players.forEach(p => {
    if (p.IsAnnounced !== "P" && p.IsAnnounced !== "S") return;

    map[p.Id] = {
      name: p.Name,
      points: Number(p.GamedayPoints || 0),
      isForeign: p.IS_FP === "1"
    };
  });
  return map;
}

function getAllGameDayIds(res){
  const now=new Date(new Date().toISOString());
  return res.Data.Value.filter(m=>parseGMTDate(m.MatchdateTime)<=now)
    .map(m=>m.TourGamedayId);
}

function getCurrentGameDayId(res){
  const now=new Date(new Date().toISOString());
  let best=null,min=1e18;
  res.Data.Value.forEach(m=>{
    const t=parseGMTDate(m.MatchdateTime);
    const diff=t-now;
    if(m.IsLive){best=m;min=0;}
    else if(diff>=0&&diff<min){min=diff;best=m;}
  });
  return best.TourGamedayId;
}

function parseGMTDate(s){
  const[d,t]=s.split(" ");
  const[m,day,y]=d.split("/");
  const[h,mi,se]=t.split(":");
  return new Date(Date.UTC(y,m-1,day,h,mi,se));
}

function fetchLeaderboard(c,g){
  return JSON.parse(UrlFetchApp.fetch(
    `${BASE_URL}/classic/api/user/leagues/6610102/leaderboard?gamedayId=${g}&optType=1&phaseId=1&pageNo=1&topNo=100&pageChunk=100&pageOneChunk=100&minCount=11&leagueId=6610102`,
    {headers:{cookie:c,entity:"d3tR0!t5m@sh"}}).getContentText());
}

function fetchTeam(c,id,s,g){
  return JSON.parse(UrlFetchApp.fetch(
    `${BASE_URL}/classic/api/user/guid/lb-team-get?gamedayId=${g}&tourgamedayId=${g}&teamId=${id}&socialId=${s}`,
    {headers:{cookie:c,entity:"d3tR0!t5m@sh"}}).getContentText());
}

function fetchPlayers(c,g){
  return JSON.parse(UrlFetchApp.fetch(
    `${BASE_URL}/classic/api/feed/gamedayplayers?lang=en&tourgamedayId=${g}&teamgamedayId=${g}`,
    {headers:{cookie:c,entity:"d3tR0!t5m@sh"}}).getContentText());
}

function fetchFixtures(c){
  return JSON.parse(UrlFetchApp.fetch(
    `${BASE_URL}/classic/api/feed/tour-fixtures?lang=en`,
    {headers:{cookie:c,entity:"d3tR0!t5m@sh"}}).getContentText());
}

function buildMatchHeader(g){
  let h=["Team Name"];
  g.forEach(x=>h.push(`G${x} C`,`G${x} VC`));
  return h;
}

function buildPlayerNameMap(pm){
  const m={};
  Object.values(pm).forEach(p=>{
    Object.entries(p).forEach(([id,v])=>m[id]=v.name);
  });
  return m;
}

function buildAverageTable(matchData, history, nameMap){
  return matchData.map(row=>{
    const name = row[0];
    let tc=0,tv=0,count=0,bc=0,bv=0,bcP="",bvP="";
    for(let i=1,mi=0;i<row.length;i+=2,mi++){
      const c=row[i]||0,vc=row[i+1]||0;
      tc+=c;tv+=vc;count++;
      const h=history[name][mi];
      if(c>bc){bc=c;bcP=nameMap[h?.c]||"";}
      if(vc>bv){bv=vc;bvP=nameMap[h?.vc]||"";}
    }
    return [
      name,
      (tc/count).toFixed(2),
      (tv/count).toFixed(2),
      ((tc+tv)/count).toFixed(2),
      `${bcP} (${bc})`,
      `${bvP} (${bv})`
    ];
  });
}

// ================= TABLE 4 =================

function buildMatchWiseTotalAndSubs(teams, gameDays, playerMaps, teamCache) {

  return teams.map(t => {

    let row = [t.temname];

    gameDays.forEach(gd => {

      const teamData = teamCache[t.temname];
      const val = teamData ? teamData[gd] : null;

      if (!val) {
        row.push(0, 0, 0, 0);
        return;
      }

      const pm = playerMaps[gd] || {};

      let basePoints = 0;
      let foreignPoints = 0;
      let indianPoints = 0;
      let playersCount = 0;

      val.plyid.forEach(pid => {

        const p = pm[pid];
        if (!p) return;

        const pts = Number(p.points || 0);

        basePoints += pts;
        playersCount++;

        if (p.isForeign) foreignPoints += pts;
        else indianPoints += pts;
      });

      // Captain / VC
      const c = pm[val.mcapt];
      const vc = pm[val.vcapt];

      const captainPoints = c ? c.points : 0;
      const vcPoints = vc ? vc.points : 0;

      let totalPoints = basePoints + captainPoints + (vcPoints * 0.5);

      // Boosters
      (val.booster || []).forEach(b => {

        if (b.cf_team_gamedayid !== gd) return;

        const before = totalPoints;

        switch (b.cf_boosterid) {
          case 3: // DOUBLE POWER
            totalPoints *= 2;
            break;

          case 9: { // FOREIGN STARS

            let extra = foreignPoints;

            if (c && c.isForeign) extra += captainPoints;
            if (vc && vc.isForeign) extra += vcPoints * 0.5;

            totalPoints += extra;
            break;
          }

          case 10: { // INDIAN WARRIORS

            let extra = indianPoints;

            if (c && !c.isForeign) extra += captainPoints;
            if (vc && !vc.isForeign) extra += vcPoints * 0.5;

            totalPoints += extra;
            break;
          }

          case 12: // TRIPLE CAPTAIN
            totalPoints += captainPoints;
            break;
        }
      });

      const ptsPerPlayer = playersCount ? totalPoints / playersCount : 0;

      row.push(
        Number(totalPoints.toFixed(1)),
        val.subgdusr || 0,
        playersCount,
        Number(ptsPerPlayer.toFixed(2))
      );
    });

    return row;
  });
}

function buildEfficiencySummary(teams, gameDays, totalData) {

  const summary = [];

  gameDays.forEach((gd, index) => {

    let bestTeam = "";
    let bestValue = -Infinity;

    let worstTeam = "";
    let worstValue = Infinity;

    totalData.forEach(row => {

      const teamName = row[0];

      // ✅ FIX: 4 columns per match
      const baseIndex = 1 + index * 4;

      const points = Number(row[baseIndex] || 0);
      const subs   = Number(row[baseIndex + 1] || 0);

      let efficiency;

      if (subs > 0) {
        efficiency = points / subs;
      } else {
        efficiency = points; // special case
      }

      // BEST
      if (efficiency > bestValue) {
        bestValue = efficiency;
        bestTeam = teamName;
      }

      // WORST
      if (efficiency < worstValue) {
        worstValue = efficiency;
        worstTeam = teamName;
      }
    });

    summary.push([
      `Game ${gd}`,
      bestTeam,
      Number(bestValue.toFixed(2)),
      worstTeam,
      Number(worstValue.toFixed(2))
    ]);
  });

  return summary;
}

function buildBalancedSubsPlan(subsLeft, matchesLeft) {

  if (matchesLeft <= 0) return "";

  const base = Math.floor(subsLeft / matchesLeft);
  let remainder = subsLeft % matchesLeft;

  const planMap = {};

  for (let i = 0; i < matchesLeft; i++) {

    let subs = base;

    if (remainder > 0) {
      subs += 1;
      remainder--;
    }

    if (subs === 0) continue;

    planMap[subs] = (planMap[subs] || 0) + 1;
  }

  // Format output
  return Object.keys(planMap)
    .sort((a,b)=>b-a)
    .map(k => `${k}x${planMap[k]}`)
    .join(" | ");
}

function buildBoosterStatsFromTable1(ctx, maps) {

  const boosterMap = {};

  ctx.teams.forEach(t => {

    const val = ctx.teamCache[t.temname]?.[ctx.gamedayId];
    if (!val) return;

    const boosterText = formatBoosterWithPoints(
      val,
      t.temname,
      maps.table4Map,
      maps.table2Map
    );

    boosterMap[t.temname] = parseBoosterText(boosterText);
  });

  return boosterMap;
}

function parseBoosterText(text) {

  const result = {
    "Wild Card": "",
    "Double Power": "",
    "Foreign Stars": "",
    "Indian Warriors": "",
    "Free Hit": "",
    "Triple Captain": ""
  };

  if (!text) return result;

  const parts = text.split(" | ");

  parts.forEach(p => {

    const [name, values] = p.split(" (");
    if (!values) return;

    const clean = values.replace(")", "");

    const entries = clean.split(",");

    let bestGame = "";
    let bestPoints = -1;

    entries.forEach(e => {
      const [g, pts] = e.split(":");
      const num = Number(pts);

      if (num > bestPoints) {
        bestPoints = num;
        bestGame = `${g.trim()}: ${num}`;
      }
    });

    result[name.trim()] = bestGame;
  });

  return result;
}