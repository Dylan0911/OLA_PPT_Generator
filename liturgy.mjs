// 礼仪日历计算：romcal v1（澳洲日历）查询 + 中文映射 + 礼仪年轮盘的季节边界。
import romcalPkg from "romcal";

const romcal = romcalPkg.default || romcalPkg;

// 季节 → 中文 + 轮盘配色
const SEASON_CN = {
  Advent: "将临期",
  Christmas: "圣诞期",
  Christmastide: "圣诞期",
  Lent: "四旬期",
  "Holy Week": "圣周",
  Easter: "复活期",
  Eastertide: "复活期",
  "Ordinary Time": "常年期",
};

// 礼仪色 → 中文 + 界面用色（白/金用金色，便于在浅色玻璃上显示）
const COLOR_CN = {
  GREEN: { cn: "绿", hex: "#4F7E63" },
  PURPLE: { cn: "紫", hex: "#5E4A9E" },
  VIOLET: { cn: "紫", hex: "#5E4A9E" },
  WHITE: { cn: "白／金", hex: "#C9A227" },
  RED: { cn: "红", hex: "#B3262A" },
  ROSE: { cn: "玫瑰", hex: "#C45B7C" },
  GOLD: { cn: "金", hex: "#C9A227" },
  BLACK: { cn: "黑", hex: "#37474F" },
};

// 季节本身的「稳定颜色」——整页与轮盘用它，所以同一节期里每天颜色不变。
const SEASON_COLOR = {
  Advent: { key: "PURPLE", cn: "紫", hex: "#5E4A9E" },
  Christmas: { key: "WHITE", cn: "白／金", hex: "#C9A227" },
  Lent: { key: "PURPLE", cn: "紫", hex: "#5E4A9E" },
  Easter: { key: "WHITE", cn: "白／金", hex: "#C9A227" },
  "Ordinary Time": { key: "GREEN", cn: "绿", hex: "#4F7E63" },
};

// 庆典等级 → 中文标签
const RANK_CN = { SOLEMNITY: "大庆日", FEAST: "庆节", MEMORIAL: "纪念", OPT_MEMORIAL: "可纪念", OPTIONAL_MEMORIAL: "可纪念" };

// 重要庆日英文名 → 中文
const FEAST_CN = [
  [/Easter Sunday|Easter Day/i, "复活节主日"],
  [/Pentecost/i, "圣神降临节"],
  [/Christ the King/i, "基督普世君王节"],
  [/Holy Trinity|Trinity Sunday/i, "天主圣三节"],
  [/Corpus Christi|Body and Blood/i, "基督圣体圣血节"],
  [/Sacred Heart/i, "耶稣圣心节"],
  [/Ascension/i, "耶稣升天节"],
  [/Ash Wednesday/i, "圣灰礼仪（四旬期首日）"],
  [/Palm Sunday|Passion Sunday/i, "圣枝主日"],
  [/Holy Thursday|Maundy Thursday/i, "主的晚餐（圣周四）"],
  [/Good Friday/i, "主受难日（圣周五）"],
  [/Holy Saturday|Easter Vigil/i, "复活前夕（圣周六）"],
  [/Mary,? Mother of God/i, "天主之母节"],
  [/Immaculate Conception/i, "圣母始胎无染原罪节"],
  [/Assumption/i, "圣母蒙召升天节"],
  [/All Saints/i, "诸圣节"],
  [/All Souls/i, "追思已亡节"],
  [/Epiphany/i, "主显节"],
  [/Baptism of the Lord/i, "主受洗节"],
  [/Annunciation/i, "圣母领报节"],
  [/Holy Family/i, "圣家节"],
  // 「诞辰」类要放在通用 Christmas 之前，避免误判
  [/(Nativity|Birth) of (Saint )?John the Baptist/i, "圣若翰洗者诞辰"],
  [/(Passion|Beheading|Martyrdom) of (Saint )?John the Baptist/i, "圣若翰洗者殉道"],
  [/Nativity of the (Blessed Virgin Mary|B\.?V\.?M)|Birth of (the )?(Blessed Virgin )?Mary/i, "圣母诞辰"],
  [/Nativity of the Lord|Christmas/i, "圣诞节"],
  [/Presentation of the Lord/i, "献主节"],
  [/Transfiguration/i, "主显圣容节"],
  [/Triumph of the Cross|Exaltation of the (Holy )?Cross|Holy Cross/i, "光荣十字圣架节"],
  [/Visitation/i, "圣母访亲节"],
  [/Queenship of (the Blessed Virgin )?Mary/i, "圣母元后纪念"],
  [/Our Lady of the Rosary/i, "玫瑰圣母纪念"],
  [/Dedication of the Lateran/i, "拉特朗大殿奉献节"],
  [/Michael,?\s*(and\s*)?Gabriel|Archangels/i, "圣弥额尔、嘉俾额尔、辣法厄尔总领天使节"],
  // 宗徒与圣史（复合名要排在单名之前）
  [/Saints? Peter and Paul/i, "圣伯多禄及圣保禄宗徒节"],
  [/Chair of (Saint )?Peter/i, "圣伯多禄宗徒建立教座节"],
  [/Conversion of (Saint )?Paul/i, "圣保禄宗徒归化节"],
  [/Saints? Philip and James/i, "圣斐理伯及圣雅各伯宗徒节"],
  [/Saints? Simon and Jude/i, "圣西满及圣犹达宗徒节"],
  [/Saint Thomas/i, "圣多默宗徒节"],
  [/Saint James/i, "圣雅各伯宗徒节"],
  [/Saint Andrew/i, "圣安德肋宗徒节"],
  [/Saint Bartholomew/i, "圣巴尔多禄茂宗徒节"],
  [/Saint Matthias/i, "圣玛弟亚宗徒节"],
  [/Saint Matthew/i, "圣玛窦宗徒圣史节"],
  [/Saint Mark/i, "圣马尔谷圣史节"],
  [/Saint Luke/i, "圣路加圣史节"],
  [/Saint John,?\s*(the )?(Apostle|Evangelist)/i, "圣若望宗徒圣史节"],
  // 其他重要庆节
  [/Saint Joseph the Worker/i, "劳工圣若瑟纪念"],
  [/Saint Joseph/i, "大圣若瑟节"],
  [/Saint Stephen/i, "圣斯德望首位殉道者节"],
  [/Holy Innocents/i, "诸圣婴孩殉道节"],
  [/Saint Lawrence/i, "圣老楞佐执事殉道节"],
  [/Saint Mary Magdalene/i, "圣玛利亚玛达肋纳节"],
  [/Mary MacKillop|Mary of (the|The) Cross/i, "圣玛利亚·麦基洛（澳洲主保）"],
  [/Saint Patrick/i, "圣巴特利爵主教节"],
];

const CN_NUM = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十", "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十", "三十一", "三十二", "三十三", "三十四"];

const SEASON_OF = { Advent: "将临期", Easter: "复活期", Lent: "四旬期", "Ordinary Time": "常年期", Christmas: "圣诞期" };

function ordinalSundayCn(name) {
  const m = name.match(/(\d+)(?:st|nd|rd|th)\s+Sunday\s+of\s+(Advent|Easter|Lent|Ordinary Time)/i);
  if (!m) return null;
  const n = Number(m[1]);
  const season = SEASON_OF[m[2].replace(/\b\w/g, (c) => c.toUpperCase()).replace("Of", "of")] || SEASON_OF[m[2]];
  return season && CN_NUM[n] ? `${season}第${CN_NUM[n]}主日` : null;
}

// 返回中文译名；没有对应翻译则返回 null（交给调用方决定显示平日还是英文）。
function nameToCnOrNull(name) {
  for (const [re, cn] of FEAST_CN) if (re.test(name)) return cn;
  const ord = ordinalSundayCn(name);
  if (ord) return ord;
  if (/Sunday/i.test(name)) return name.replace(/Sunday/i, "主日");
  return null;
}

// ── 礼仪年轮盘：用复活节算法 + 将临期公式算出各季节边界（用于可视化） ──
function easterSunday(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return Date.UTC(year, month - 1, day);
}

const DAY = 86400000;
const addDays = (t, n) => t + n * DAY;

function adventStart(year) {
  const xmas = Date.UTC(year, 11, 25);
  const dow = new Date(xmas).getUTCDay();
  const advent4 = addDays(xmas, -(dow === 0 ? 7 : dow));
  return addDays(advent4, -21);
}

function sundayAfter(t) {
  const dow = new Date(t).getUTCDay();
  return addDays(t, dow === 0 ? 7 : 7 - dow);
}

// 返回所选日期所在「礼仪年」的季节分段（占比）+ 标记位置。
function yearRing(dateUTC) {
  const y = new Date(dateUTC).getUTCFullYear();
  const startYear = dateUTC >= adventStart(y) ? y : y - 1;
  const start = adventStart(startYear);
  const end = adventStart(startYear + 1);
  const nextYear = startYear + 1;
  const easter = easterSunday(nextYear);
  const ashWed = addDays(easter, -46);
  const pentecost = addDays(easter, 49);
  const baptism = sundayAfter(Date.UTC(nextYear, 0, 6)); // 主受洗节（约略）
  const span = end - start;
  const frac = (t) => Math.max(0, Math.min(1, (t - start) / span));
  const segs = [
    { key: "Advent", cn: "将临期", hex: "#5E4A9E", a: start, b: Date.UTC(startYear, 11, 25) },
    { key: "Christmas", cn: "圣诞期", hex: "#C9A227", a: Date.UTC(startYear, 11, 25), b: baptism },
    { key: "OrdinaryA", cn: "常年期", hex: "#4F7E63", a: baptism, b: ashWed },
    { key: "Lent", cn: "四旬期", hex: "#5E4A9E", a: ashWed, b: easter },
    { key: "Easter", cn: "复活期", hex: "#C9A227", a: easter, b: addDays(pentecost, 1) },
    { key: "OrdinaryB", cn: "常年期", hex: "#4F7E63", a: addDays(pentecost, 1), b: end },
  ].map((s) => ({ key: s.key, cn: s.cn, hex: s.hex, start: frac(s.a), end: frac(s.b) }));
  return { segments: segs, marker: frac(dateUTC) };
}

const calCache = new Map();
function calendarForYear(year) {
  if (!calCache.has(year)) {
    calCache.set(
      year,
      romcal.calendarFor({ year, country: "australia", locale: "en", ascensionOnSunday: true })
    );
  }
  return calCache.get(year);
}

const isoOf = (m) => new Date(m).toISOString().slice(0, 10);

export async function liturgyForDate(dateISO) {
  const clean = /^\d{4}-\d{2}-\d{2}$/.test(dateISO) ? dateISO : new Date().toISOString().slice(0, 10);
  const dateUTC = Date.UTC(...clean.split("-").map((n, i) => (i === 1 ? Number(n) - 1 : Number(n))));
  const year = Number(clean.slice(0, 4));
  const cal = await calendarForYear(year);
  const entries = cal.filter((d) => isoOf(d.moment) === clean);
  // 同一天可能有多条（主日 + 可选纪念）；取等级最高的（prioritized / 类型）
  const rank = { SOLEMNITY: 5, FEAST: 4, SUNDAY: 3, MEMORIAL: 2 };
  const day =
    entries.sort((a, b) => (b.data?.prioritized ? 1 : 0) - (a.data?.prioritized ? 1 : 0) || (rank[b.type] || 1) - (rank[a.type] || 1))[0] ||
    entries[0];

  if (!day) return { date: clean, error: "no-data" };

  const seasonKey = day.data?.season?.key || "";
  const seasonVal = day.data?.season?.value || "";
  const weekday = new Date(dateUTC).getUTCDay();
  const isSunday = weekday === 0;
  const highlight = day.type === "SOLEMNITY" || day.type === "FEAST";

  // 季节色：整页主色 + 轮盘用它 —— 同一节期里每天稳定不变。
  const season = SEASON_COLOR[seasonKey] || SEASON_COLOR[seasonVal] || { key: "GREEN", cn: "绿", hex: "#4F7E63" };

  // 当日「祭衣色」：主日/大庆日/庆节用当天的真实礼仪色；平日与（可）纪念沿用季节色，避免天天乱跳。
  const romcalKey = day.data?.meta?.liturgicalColor?.key || season.key;
  const romcalColor = COLOR_CN[romcalKey] || { cn: romcalKey, hex: "#4F7E63" };
  const vestment = isSunday || highlight ? { key: romcalKey, ...romcalColor } : { key: season.key, cn: season.cn, hex: season.hex };

  // 名称：主日/大庆日/庆节及有中文译名者 → 译名；其余平日 → 「<节期>平日」，圣人放小字备注。
  const translated = nameToCnOrNull(day.name);
  const seasonCn = SEASON_CN[seasonKey] || SEASON_CN[seasonVal] || seasonVal;
  const useProperName = translated || isSunday || highlight;
  const nameCn = useProperName ? translated || day.name : `${seasonCn}平日`;
  const nameEn = useProperName ? day.name : "";
  const memo = useProperName ? "" : day.name; // 平日里的（可）纪念圣人，作小字

  return {
    date: clean,
    weekdayCn: ["主日", "周一", "周二", "周三", "周四", "周五", "周六"][weekday],
    nameCn,
    nameEn,
    memo,
    type: day.type,
    rankCn: RANK_CN[day.type] || "",
    seasonCn,
    seasonEn: seasonVal,
    seasonHex: season.hex,
    // 当日祭衣色（礼仪色卡片 + 生成 PPT 用色 + 轮盘标记）
    colorKey: vestment.key,
    colorCn: vestment.cn,
    colorHex: vestment.hex,
    // 与季节色是否不同（用于高亮「今日庆日」）
    special: vestment.hex.toUpperCase() !== season.hex.toUpperCase(),
    cycle: day.data?.meta?.cycle?.value || "",
    cycleCn: { A: "甲年", B: "乙年", C: "丙年" }[(day.data?.meta?.cycle?.value || "").slice(-1)] || "",
    psalterWeek: typeof day.data?.meta?.psalterWeek?.key === "number" ? day.data.meta.psalterWeek.key : "",
    ring: yearRing(dateUTC),
  };
}
