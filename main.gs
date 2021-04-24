const main = () => {
  const contestInfo = crawlUpcomingContests();
  const newContests = updateSpreadsheet(contestInfo);
  if (newContests.length) Logger.log(newContests);
  const formattedPostData = formatData(newContests);
  formattedPostData.forEach((d) => {
    updateCalender(d[0], d[1], d[2], d[3]);
  });
};

const crawlUpcomingContests = () => {
  const url = "https://atcoder.jp/contests/";
  const postheader = {
    accept: "gzip, */*",
    timeout: "20000",
  };

  const parameters = {
    method: "get",
    muteHttpExceptions: true,
    headers: postheader,
  };

  const content = UrlFetchApp.fetch(url, parameters).getContentText("UTF-8");
  const upcomingContests = content
    .match(
      /<div id="contest-table-upcoming">[\S\s]*<hr>[\s]*<div id="contest-table-recent">/
    )[0]
    .match(/<tr>[\s\S]*?<\/tr>/g);

  const extractInfo = (trHtml) => {
    const tds = trHtml.match(/<td[\s\S]*?<\/td>/g);

    const time = tds[0].match(/iso=\d{8}T\d{4}/)[0].replace("iso=", "");
    const contestName = tds[1]
      .match(/<a[\s\S]*?<\/a>/)[0]
      .replace(/<[\s\S]*?>/g, "");
    const duration = tds[2]
      .match(/<td[\s\S]*?<\/td>/)[0]
      .replace(/<[\s\S]*?>/g, "");
    const rated = tds[3]
      .match(/<td[\s\S]*?<\/td>/)[0]
      .replace(/<[\s\S]*?>/g, "");

    return {
      time: time,
      contestName: contestName,
      duration: duration,
      rated: rated,
    };
  };

  //一つ目は表のヘッダなので省く
  const contestInfo = upcomingContests.slice(1).map((d) => extractInfo(d));

  return contestInfo;
};

const updateSpreadsheet = (data) => {
  const spreadsheet = SpreadsheetApp.getActiveSheet();

  let lastRow = spreadsheet.getLastRow();
  const contestNames = spreadsheet.getRange(1, 2, lastRow).getValues();

  const newContests = [];

  data.forEach((d, i) => {
    //すでにスプレッドシートに登録済みならパス
    for (let j = lastRow - 1; j >= 1; j--) {
      if (d.contestName === contestNames[j][0]) {
        return;
      }
    }
    //新しかったらシートに記録
    spreadsheet
      .getRange(lastRow + i + 1, 1, 1, 4)
      .setValues([[d.time, d.contestName, d.duration, d.rated]]);
    newContests.push([d.time, d.contestName, d.duration, d.rated]);
  });
  return newContests;
};

const formatData = (data) => {
  const toRFC3339Format = (time) => {
    const date = time.toISOString().slice(0, 11);
    let hour = time.getHours();
    if (hour < 10) {
      hour = "0" + hour;
    }
    let minute = time.getMinutes();
    if (minute < 10) {
      minute = "0" + minute;
    }
    return date + hour + ":" + minute + ":00";
  };

  const formattedPostData = data.map((d) => {
    const year = Number(d[0].slice(0, 4));
    const month = Number(d[0].slice(4, 6)) - 1;
    const date = Number(d[0].slice(6, 8)) - 1;
    const hour = Number(d[0].slice(-4, -2));

    const firstColon = d[2].indexOf(":");
    const durationHour = Number(d[2].slice(0, firstColon));
    const durationMinute = Number(d[2].slice(firstColon + 1, firstColon + 3));

    const start = new Date(year, month, date, hour);
    const end = new Date(year, month, date, hour);

    end.setHours(end.getHours() + durationHour);
    end.setMinutes(end.getMinutes() + durationMinute);

    return [toRFC3339Format(start), toRFC3339Format(end), d[1], d[3]];
  });

  return formattedPostData;
};

const updateCalender = (start, end, contestName, rated) => {
  const ServiceAccountKey = JSON.parse(
    PropertiesService.getScriptProperties().getProperty("serviceAccountJsonKey")
  );

  const calendarId = PropertiesService.getScriptProperties().getProperty(
    "calenderID"
  );

  const getService = () => {
    return OAuth2.createService("calendar")
      .setAuthorizationBaseUrl("https://accounts.google.com/o/oauth2/auth")
      .setTokenUrl("https://accounts.google.com/o/oauth2/token")
      .setPrivateKey(ServiceAccountKey.private_key)
      .setIssuer(ServiceAccountKey.client_email)
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope("https://www.googleapis.com/auth/calendar");
  };

  const payload = {
    start: {
      dateTime: start,
      timeZone: "Asia/Tokyo",
    },
    end: {
      dateTime: end,
      timeZone: "Asia/Tokyo",
    },
    summary: contestName,
    description: "rated: " + rated,
  };

  const service = getService();

  const fetchOptions = {
    method: "post",
    payload: JSON.stringify(payload),
    contentType: "application/json",
    headers: { Authorization: "Bearer " + service.getAccessToken() },
  };

  const url =
    "https://www.googleapis.com/calendar/v3/calendars/" +
    calendarId +
    "/events";

  UrlFetchApp.fetch(url, fetchOptions);
};
