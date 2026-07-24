/**
 * ORDER TRACKER — Apps Script 后端 (v2 · 名字优先)
 * ------------------------------------------------------------
 * 数据存在同一个 Google Sheet 里，4 个分页：
 *   Orders  - 每一笔订单
 *   Config  - salesman → branch / state / region 名单 (全年 Excel 去重，285 人)
 *   Rates   - 抽佣规则 (按地区，可在 App 里改)
 *   Prices  - 每个 branch 各 set 类型的默认价格 (录单时自动带出，可改)
 * 部署成网页 App 后，合伙人打开同一链接即可共用同一份数据。
 * ------------------------------------------------------------
 * 第一次使用：在编辑器里手动运行 setupSpreadsheet() 一次做初始化。
 */

var SHEET_ORDERS = 'Orders';
var SHEET_CONFIG = 'Config';
var SHEET_RATES = 'Rates';
var SHEET_PRICES = 'Prices';

var ORDERS_HEADERS = ['ID', 'Timestamp', 'Date', 'Region', 'State', 'Branch', 'Salesman',
  'Item', 'SetDetail', 'Price', 'Set', 'Total', 'Commission', 'Status', 'PaymentDate', 'Month', 'EnteredBy'];
var CONFIG_HEADERS = ['Region', 'State', 'Branch', 'Salesman'];
var RATES_HEADERS = ['Region', 'Item', 'Mode', 'Value', 'Note'];
var PRICES_HEADERS = ['Branch', 'SetType', 'Price'];

var REGION_LABELS = { kl: 'KL / Selangor', ns: 'Seremban / NS', johor: 'Johor' };

// ================= 种子数据 (来自客户全年 Excel，已去重规范化) =================

var SEED_CONFIG = [
  ['johor','Johor','Ang Trading','SHALEN'],
  ['johor','Johor','Chery Skudai','ROGEL'],
  ['johor','Johor','Honda Tebrau','EZAT'],
  ['johor','Johor','Honda Tebrau','SHAHMAN'],
  ['johor','Johor','Jaecoo Tmn Daya','AGNES'],
  ['johor','Johor','Jaecoo Tmn Daya','CALYN'],
  ['johor','Johor','Jaecoo Tmn Daya','CK'],
  ['johor','Johor','Jaecoo Tmn Daya','ELIZ'],
  ['johor','Johor','Jaecoo Tmn Daya','HARRY'],
  ['johor','Johor','Jaecoo Tmn Daya','JOSEPHINE'],
  ['johor','Johor','Jaecoo Tmn Daya','LUCAS'],
  ['johor','Johor','Jaecoo Tmn Daya','PARKER'],
  ['johor','Johor','Jaecoo Tmn Daya','SAMANTHA'],
  ['johor','Johor','Jaecoo Tmn Daya','STEPHANIE'],
  ['johor','Johor','Jaecoo Tmn Daya','THOMAS'],
  ['johor','Johor','Jaecoo Tmn Daya','VICKY'],
  ['johor','Johor','Jetour Kebun Teh','DENNIS'],
  ['johor','Johor','Jetour Kebun Teh','WALTON'],
  ['johor','Johor','Jetour Kebun Teh','WATON'],
  ['johor','Johor','Mit Skudai','FIEMA'],
  ['johor','Johor','Mitsubishi Skudai','FIMA'],
  ['johor','Johor','Perodua Bkt Gambir','AMIRA'],
  ['johor','Johor','Perodua Bkt Gambir','AMNIE'],
  ['johor','Johor','Perodua Bkt Gambir','AMRY'],
  ['johor','Johor','Perodua Bkt Gambir','ASIKIN'],
  ['johor','Johor','Perodua Bkt Gambir','FADHLI'],
  ['johor','Johor','Perodua Bkt Gambir','FAI'],
  ['johor','Johor','Perodua Bkt Gambir','FIE'],
  ['johor','Johor','Perodua Bkt Gambir','HAFIZAN'],
  ['johor','Johor','Perodua Bkt Gambir','HARISMAN'],
  ['johor','Johor','Perodua Bkt Gambir','HAZMI'],
  ['johor','Johor','Perodua Bkt Gambir','ISA'],
  ['johor','Johor','Perodua Bkt Gambir','MEL'],
  ['johor','Johor','Perodua Bkt Gambir','NAZRI'],
  ['johor','Johor','Perodua Kluang','CARWAN'],
  ['johor','Johor','Perodua Kluang','FAREZ'],
  ['johor','Johor','Perodua Kluang','MS.TAN'],
  ['johor','Johor','Perodua Kluang','SABREE'],
  ['johor','Johor','Perodua Kota Masai','AZIE'],
  ['johor','Johor','Perodua Kota Masai','AZWA'],
  ['johor','Johor','Perodua Kota Masai','DAUS'],
  ['johor','Johor','Perodua Kota Masai','EPUL'],
  ['johor','Johor','Perodua Kota Masai','FARA'],
  ['johor','Johor','Perodua Kota Masai','FIKRI'],
  ['johor','Johor','Perodua Kota Masai','HAN'],
  ['johor','Johor','Perodua Kota Masai','HUSAINI'],
  ['johor','Johor','Perodua Kota Masai','IJAM'],
  ['johor','Johor','Perodua Kota Masai','IZZAT'],
  ['johor','Johor','Perodua Kota Masai','LOTFI'],
  ['johor','Johor','Perodua Kota Masai','OMAR'],
  ['johor','Johor','Perodua Kota Masai','VAMP'],
  ['johor','Johor','Perodua Kota Masai','YAN'],
  ['johor','Johor','Perodua Kota Masai','YATIE'],
  ['johor','Johor','Perodua Kota Masai','YIBBU'],
  ['johor','Johor','Perodua Kota Masai','ZIE'],
  ['johor','Johor','Perodua Kota Masai','ZULAIKHA'],
  ['johor','Johor','Perodua Muar','ANUAR'],
  ['johor','Johor','Perodua Muar','BAQRIN'],
  ['johor','Johor','Perodua Muar','FOO'],
  ['johor','Johor','Perodua Muar','HAKIM'],
  ['johor','Johor','Perodua Muar','HARITH'],
  ['johor','Johor','Perodua Muar','HASSAN'],
  ['johor','Johor','Perodua Muar','LIM'],
  ['johor','Johor','Perodua Muar','MARIANA'],
  ['johor','Johor','Perodua Muar','MIZI'],
  ['johor','Johor','Perodua Muar','NADIAH'],
  ['johor','Johor','Perodua Muar','SAHLEE'],
  ['johor','Johor','Perodua Muar','SYAHMAN'],
  ['johor','Johor','Perodua Muar','TARMIZI'],
  ['johor','Johor','Perodua Tebrau','KAREN'],
  ['johor','Johor','Proton Kesang','ANNE'],
  ['johor','Johor','Proton Kesang','AYUB'],
  ['johor','Johor','Proton Kesang','EDHAR'],
  ['johor','Johor','Proton Kesang','IJAL'],
  ['johor','Johor','Proton Kesang','LIZA'],
  ['johor','Johor','Proton Kesang','SHAH'],
  ['johor','Johor','Proton Kesang','YANM'],
  ['johor','Johor','Proton Larkin','FADHLIYANA'],
  ['johor','Johor','Proton Larkin','FUD'],
  ['johor','Johor','Proton Pasir Gudang','ADDIN'],
  ['johor','Johor','Proton Pasir Gudang','AISYAH'],
  ['johor','Johor','Proton Pasir Gudang','DEE'],
  ['johor','Johor','Proton Pasir Gudang','ERIN'],
  ['johor','Johor','Proton Pasir Gudang','FIERA'],
  ['johor','Johor','Proton Pasir Gudang','IRA'],
  ['johor','Johor','Proton Pasir Gudang','IZA'],
  ['johor','Johor','Proton Pasir Gudang','IZZA'],
  ['johor','Johor','Proton Pasir Gudang','MALIK'],
  ['johor','Johor','Proton Pasir Gudang','MUS'],
  ['johor','Johor','Proton Pasir Gudang','MUSTAQIM'],
  ['johor','Johor','Proton Pasir Gudang','QISTINA'],
  ['johor','Johor','Proton Pasir Gudang','SAIDI'],
  ['johor','Johor','Proton Pasir Gudang','SITI'],
  ['johor','Johor','Proton Pasir Gudang','TASHA'],
  ['johor','Johor','Proton Pasir Gudang','WARI'],
  ['johor','Johor','Proton Pasir Gudang','YANA'],
  ['johor','Johor','Proton Pasir Gudang','ZUL'],
  ['johor','Johor','Proton Pasir Gudang','ZUNIE'],
  ['johor','Johor','Proton Pasir Gudang','ZUNNIE'],
  ['johor','Johor','Proton Skudai','FIRDAUS'],
  ['johor','Johor','Proton Skudai','TERENCE'],
  ['johor','Johor','Proton Skudai','TERENCE LEE'],
  ['johor','Johor','Proton Tebrau','BETTY'],
  ['johor','Johor','eMAS','FREDDIE'],
  ['johor','Johor','eMAS Pasir Gudang','NAZZURIN'],
  ['johor','Johor','eMAS Pasir Gudang','RYAN'],
  ['kl','Selangor','Chery Balakong','LOUIS'],
  ['kl','Selangor','Chery Bangi','AWI'],
  ['kl','Selangor','Chery Bangi','MJ'],
  ['kl','Selangor','Chery Cheras','JACK TAN'],
  ['kl','Selangor','Chery Connaught','JACK'],
  ['kl','Selangor','Chery Nilai','CHOW'],
  ['kl','Selangor','Chery Semenyih','JIM'],
  ['kl','Selangor','Chery Senawang','TAN'],
  ['kl','Selangor','E Mas Putrajaya','SHAFIQ'],
  ['kl','Selangor','Honda Cyberjaya','AFZAL'],
  ['kl','Selangor','Honda Cyberjaya','AINI'],
  ['kl','Selangor','Honda Cyberjaya','ANN'],
  ['kl','Selangor','Honda Cyberjaya','AYU'],
  ['kl','Selangor','Honda Cyberjaya','AZIM'],
  ['kl','Selangor','Honda Cyberjaya','AZZAM'],
  ['kl','Selangor','Honda Cyberjaya','BIYA'],
  ['kl','Selangor','Honda Cyberjaya','BIYAA'],
  ['kl','Selangor','Honda Cyberjaya','FARAH'],
  ['kl','Selangor','Honda Cyberjaya','HANA'],
  ['kl','Selangor','Honda Cyberjaya','INTAN'],
  ['kl','Selangor','Honda Cyberjaya','IZZ'],
  ['kl','Selangor','Honda Cyberjaya','JA'],
  ['kl','Selangor','Honda Cyberjaya','JEMAN'],
  ['kl','Selangor','Honda Cyberjaya','MIRA'],
  ['kl','Selangor','Honda Cyberjaya','MUAAZ'],
  ['kl','Selangor','Honda Cyberjaya','ROS'],
  ['kl','Selangor','Honda Cyberjaya','ROSE'],
  ['kl','Selangor','Honda Cyberjaya','ZOOL'],
  ['kl','Selangor','Honda Seri Kembangan','KELVIN'],
  ['kl','Selangor','Jaecoo Semenyih','KEAVIN'],
  ['kl','Selangor','Jaecoo Semenyih','OOI LI PING'],
  ['kl','Selangor','Jetour Putrajaya','KUMAR'],
  ['kl','Selangor','Mitsubishi Bangi','WEE'],
  ['kl','Selangor','Mitsubishi Jln Ampang','FARHANIS'],
  ['kl','Selangor','Mitubushi Balakong','ZIKRI'],
  ['kl','Selangor','Perodua Bangi','AMIN'],
  ['kl','Selangor','Perodua Bangi','AMINUDDIN'],
  ['kl','Selangor','Perodua Bangi','AMINUDIN'],
  ['kl','Selangor','Perodua Bangi','ASYRAF'],
  ['kl','Selangor','Perodua Bangi','KHAIRUL'],
  ['kl','Selangor','Perodua Bangi','TAUFIK'],
  ['kl','Selangor','Perodua Kajang','LISA'],
  ['kl','Selangor','Perodua Kajang','MIELA'],
  ['kl','Selangor','Perodua Kajang Prima','ADIBAH'],
  ['kl','Selangor','Perodua Kajang Prima','ANI'],
  ['kl','Selangor','Perodua Kajang Prima','ANIS'],
  ['kl','Selangor','Perodua Kajang Prima','AZILA'],
  ['kl','Selangor','Perodua Kajang Prima','AZMI'],
  ['kl','Selangor','Perodua Kajang Prima','IDA'],
  ['kl','Selangor','Perodua Kajang Prima','LIN'],
  ['kl','Selangor','Perodua Kajang Prima','MAZLI'],
  ['kl','Selangor','Perodua Kajang Prima','RUBY'],
  ['kl','Selangor','Perodua Kajang Prima','SUE'],
  ['kl','Selangor','Perodua Kajang Prima','SYAFIQ'],
  ['kl','Selangor','Perodua Kajang Prima','VIRA'],
  ['kl','Selangor','Perodua Kajang Prima','WANI'],
  ['kl','Selangor','Perodua Kota Warisan','FARID'],
  ['kl','Selangor','Perodua Nilai','ZUVIRA'],
  ['kl','Selangor','Proton Balakong','AMIL'],
  ['kl','Selangor','Proton Bangi','ADLY'],
  ['kl','Selangor','Proton Bangi','AFIF'],
  ['kl','Selangor','Proton Bangi','DALI'],
  ['kl','Selangor','Proton Bangi','DELI'],
  ['kl','Selangor','Proton Bangi','JUN'],
  ['kl','Selangor','Proton Bangi','MIRUL'],
  ['kl','Selangor','Proton Beranang','SARA'],
  ['kl','Selangor','Proton Kajang','NASRUL'],
  ['kl','Selangor','Proton Kota Warisan','JEFFREY'],
  ['kl','Selangor','Proton Kota Warisan','JEFFRI'],
  ['kl','Selangor','Proton Putrajaya','AMMAR'],
  ['kl','Selangor','Proton Putrajaya','AZHAN'],
  ['kl','Selangor','Proton Putrajaya','AZIEN'],
  ['kl','Selangor','Proton Putrajaya','AZLAN'],
  ['kl','Selangor','Proton Putrajaya','AZREEN'],
  ['kl','Selangor','Proton Putrajaya','AZRIEN'],
  ['kl','Selangor','Proton Putrajaya','AZRIN'],
  ['kl','Selangor','Proton Putrajaya','GURMESH'],
  ['kl','Selangor','Proton Putrajaya','HAFEEZ'],
  ['kl','Selangor','Proton Putrajaya','HAIKAL'],
  ['kl','Selangor','Proton Putrajaya','HAKIMIE'],
  ['kl','Selangor','Proton Putrajaya','IVAN'],
  ['kl','Selangor','Proton Putrajaya','JOJEE'],
  ['kl','Selangor','Proton Putrajaya','KIMIE'],
  ['kl','Selangor','Proton Putrajaya','MAD'],
  ['kl','Selangor','Proton Putrajaya','MARCUS'],
  ['kl','Selangor','Proton Putrajaya','MUIZ'],
  ['kl','Selangor','Proton Putrajaya','MUIZZ'],
  ['kl','Selangor','Proton Putrajaya','NAZMI'],
  ['kl','Selangor','Proton Putrajaya','SAIDAH'],
  ['kl','Selangor','Proton Putrajaya','SAIFUDDIN'],
  ['kl','Selangor','Proton Putrajaya','SAIFUL'],
  ['kl','Selangor','Proton Putrajaya','SYUE'],
  ['kl','Selangor','Proton Putrajaya','ZIELA'],
  ['kl','Selangor','Proton Semenyih','EVELYN SIEW'],
  ['kl','Selangor','Proton Semenyih','VERON'],
  ['kl','Selangor','Proton Semenyih','WAYKUAN'],
  ['kl','Selangor','Proton Seri Kembangan','ABDULLAH'],
  ['kl','Selangor','Proton Seri Kembangan','ABDULLAH 1'],
  ['kl','Selangor','Proton Seri Kembangan','SAM YIP'],
  ['kl','Selangor','Proton Sg Long','DE LUNG'],
  ['kl','Selangor','Proton Sg Long','GWEN'],
  ['kl','Selangor','Proton Sg Long','JAGATHIES'],
  ['kl','Selangor','Proton Sg Long','JEGA'],
  ['kl','Selangor','Proton Sg Long','NAS'],
  ['kl','Selangor','Proton Sg Long','NAZLI'],
  ['kl','Selangor','Proton Sg Long','SHAWN'],
  ['kl','Selangor','Proton Sg Long','SYAWAN'],
  ['kl','Selangor','Proton Sg Long','SYAZWAN'],
  ['kl','Selangor','eMAS Proton Putrajaya','AISHA'],
  ['kl','Selangor','eMAS Proton Putrajaya','BRANLOH'],
  ['kl','Selangor','eMAS Proton Putrajaya','JOAN'],
  ['kl','Selangor','eMAS Proton Putrajaya','RIN'],
  ['kl','Selangor','eMAS Proton Putrajaya','SHIEYLA'],
  ['kl','Selangor','eMAS Putrajaya','ADHWA'],
  ['kl','Selangor','eMAS Putrajaya','ADI'],
  ['kl','Selangor','eMAS Putrajaya','AIZARD'],
  ['kl','Selangor','eMAS Putrajaya','DANISH'],
  ['kl','Selangor','eMAS Putrajaya','DHELA'],
  ['kl','Selangor','eMAS Putrajaya','FAYYADH'],
  ['kl','Selangor','eMAS Putrajaya','IVY'],
  ['kl','Selangor','eMAS Putrajaya','JENSEN'],
  ['kl','Selangor','eMAS Putrajaya','PIERRE'],
  ['kl','Selangor','eMAS Putrajaya','QILA'],
  ['kl','Selangor','eMAS Putrajaya','SARAH'],
  ['kl','Selangor','eMAS Putrajaya','SYAUQI'],
  ['kl','Selangor','eMAS Putrajaya','WANIE'],
  ['kl','Selangor','eMAS Putrajaya','YING'],
  ['ns','NS','Chery Nilai','HAJARUL'],
  ['ns','NS','Chery Nilai','MR CHOW'],
  ['ns','NS','Chery Senawang','MS TAN'],
  ['ns','NS','Chery Senawang','SYED'],
  ['ns','NS','Honda Nilai Impian','AMIRUL'],
  ['ns','NS','Honda Nilai Impian','HANIS'],
  ['ns','NS','Honda Nilai Impian','ZAIDA'],
  ['ns','NS','Jaecoo Seremban 2','MS. REENA'],
  ['ns','NS','Jetour Seremban','EKHSAN'],
  ['ns','NS','Mitsubishi Nilai','JEREMY'],
  ['ns','NS','Perodua Kajang Prima','AZNI'],
  ['ns','NS','Perodua Nilai','ALIF'],
  ['ns','NS','Perodua Nilai','ALIN'],
  ['ns','NS','Perodua Nilai','ARIENA'],
  ['ns','NS','Perodua Nilai','AZRUL'],
  ['ns','NS','Perodua Nilai','DANIEL'],
  ['ns','NS','Perodua Nilai','DENNESH'],
  ['ns','NS','Perodua Nilai','FAIZAL'],
  ['ns','NS','Perodua Nilai','FLORA'],
  ['ns','NS','Perodua Nilai','GUNA'],
  ['ns','NS','Perodua Nilai','HAZRI'],
  ['ns','NS','Perodua Nilai','ISMAN'],
  ['ns','NS','Perodua Nilai','KAMIL'],
  ['ns','NS','Perodua Nilai','MORGAN'],
  ['ns','NS','Perodua Nilai','NIZAM'],
  ['ns','NS','Perodua Nilai','NURUL'],
  ['ns','NS','Perodua Nilai','SELAMAT'],
  ['ns','NS','Perodua Nilai','SHUHAIMI'],
  ['ns','NS','Perodua Nilai','SHUKRI'],
  ['ns','NS','Perodua Nilai','WAN'],
  ['ns','NS','Perodua Rasah Jaya','AIDIL'],
  ['ns','NS','Perodua Rasah Jaya','AISHAH'],
  ['ns','NS','Perodua Rasah Jaya','ANDY'],
  ['ns','NS','Perodua Rasah Jaya','FAZIL'],
  ['ns','NS','Perodua Rasah Jaya','FAZIR'],
  ['ns','NS','Perodua Rasah Jaya','LIZZA'],
  ['ns','NS','Perodua Rasah Jaya','SABRINA'],
  ['ns','NS','Perodua Rasah Jaya','SILVA'],
  ['ns','NS','Perodua Rasah Jaya','SIVA'],
  ['ns','NS','Perodua Rasah Jaya','VANI'],
  ['ns','NS','Perodua Seremban2','AHMED'],
  ['ns','NS','Perodua Seremban2','ASHEK'],
  ['ns','NS','Perodua Seremban2','DANUS'],
  ['ns','NS','Perodua Seremban2','FADZIL'],
  ['ns','NS','Perodua Seremban2','HALIM'],
  ['ns','NS','Perodua Seremban2','HAZWAN'],
  ['ns','NS','Perodua Seremban2','INA'],
  ['ns','NS','Perodua Seremban2','LIEZA'],
  ['ns','NS','Perodua Seremban2','NAZRIN'],
  ['ns','NS','Perodua Tampin','NOREEN'],
  ['ns','NS','Proton Seremban','EMMA'],
  ['ns','NS','Proton Seremban2','REENA']
];

var SEED_PRICES = {
  'Ang Trading': {'UMBRELLA':25},
  'Chery Balakong': {'UMBRELLA':20},
  'Chery Bangi': {'10 ITEMS':55},
  'Chery Cheras': {'10 ITEMS':55,'8 ITEMS':55,'UMBRELLA':20},
  'Chery Connaught': {'8 ITEMS':55},
  'Chery Nilai': {'10 ITEMS':55,'8 ITEMS':55},
  'Chery Semenyih': {'10 ITEMS':55},
  'Chery Senawang': {'10 ITEMS':55},
  'Chery Skudai': {'10 ITEMS':55,'8 ITEMS':55},
  'E Mas Putrajaya': {'10 ITEMS':55,'8 ITEMS':55},
  'Honda Cyberjaya': {'10 ITEMS':50},
  'Honda Nilai Impian': {'10 ITEMS':50},
  'Honda Seri Kembangan': {'10 ITEMS':50},
  'Honda Tebrau': {'10 ITEMS':55,'9 ITEMS':40},
  'Jaecoo Semenyih': {'8 ITEMS':55,'UMBRELLA':20},
  'Jaecoo Seremban 2': {'10 ITEMS':50,'8 ITEMS':60},
  'Jaecoo Tmn Daya': {'8 ITEMS':55,'UMBRELLA':26},
  'Jetour Kebun Teh': {'8 ITEMS':55,'UMBRELLA':26},
  'Jetour Putrajaya': {'8 ITEMS':55},
  'Jetour Seremban': {'10 ITEMS':55,'8 ITEMS':55},
  'Mitsubishi Bangi': {'10 ITEMS':55,'8 ITEMS':55,'UMBRELLA':20},
  'Mitsubishi Jln Ampang': {'8 ITEMS':55},
  'Mitsubishi Nilai': {'UMBRELLA':16},
  'Mitsubishi Skudai': {'10 ITEMS':55},
  'Mitubushi Balakong': {'10 ITEMS':55},
  'Perodua Bangi': {'10 ITEMS':50,'9 ITEMS':40,'UMBRELLA':20},
  'Perodua Bkt Gambir': {'9 ITEMS':38,'UMBRELLA':1},
  'Perodua Kajang': {'9 ITEMS':35},
  'Perodua Kajang Prima': {'10 ITEMS':50,'9 ITEMS':35},
  'Perodua Kluang': {'10 ITEMS':45,'9 ITEMS':38,'UMBRELLA':13},
  'Perodua Kota Masai': {'9 ITEMS':38,'UMBRELLA':14},
  'Perodua Kota Warisan': {'10 ITEMS':45,'9 ITEMS':42,'UMBRELLA':12},
  'Perodua Muar': {'10 ITEMS':50,'UMBRELLA':13},
  'Perodua Nilai': {'10 ITEMS':40,'9 ITEMS':35,'UMBRELLA':12},
  'Perodua Rasah Jaya': {'9 ITEMS':35,'UMBRELLA':16},
  'Perodua Seremban2': {'10 ITEMS':50,'9 ITEMS':35,'UMBRELLA':12},
  'Perodua Tampin': {'UMBRELLA':11},
  'Perodua Tebrau': {'UMBRELLA':15},
  'Proton Balakong': {'10 ITEMS':50},
  'Proton Bangi': {'10 ITEMS':50,'9 ITEMS':50,'UMBRELLA':25},
  'Proton Beranang': {'10 ITEMS':45},
  'Proton Kajang': {'10 ITEMS':55,'9 ITEMS':40},
  'Proton Kesang': {'9 ITEMS':38,'UMBRELLA':12},
  'Proton Larkin': {'10 ITEMS':45,'UMBRELLA':12},
  'Proton Pasir Gudang': {'10 ITEMS':50,'9 ITEMS':40,'UMBRELLA':20},
  'Proton Putrajaya': {'10 ITEMS':50,'UMBRELLA':20},
  'Proton Semenyih': {'10 ITEMS':45,'8 ITEMS':55},
  'Proton Seremban': {'9 ITEMS':40},
  'Proton Seremban2': {'10 ITEMS':50},
  'Proton Seri Kembangan': {'10 ITEMS':50,'8 ITEMS':55,'9 ITEMS':40,'UMBRELLA':20},
  'Proton Sg Long': {'10 ITEMS':50,'UMBRELLA':11},
  'Proton Skudai': {'10 ITEMS':50},
  'Proton Tebrau': {'10 ITEMS':50},
  'eMAS': {'8 ITEMS':55},
  'eMAS Pasir Gudang': {'8 ITEMS':55},
  'eMAS Proton Putrajaya': {'10 ITEMS':50,'8 ITEMS':55,'UMBRELLA':20},
  'eMAS Putrajaya': {'10 ITEMS':55,'8 ITEMS':55,'UMBRELLA':20}
};

// 抽佣规则（可在 App 管理页调整）
//  kl/ns/johor 三个地区 + umbrella 特例
var SEED_RATES = [
  ['kl', 'standard', 'fixed_per_set', 2.5, 'KL/Selangor: RM2.50 x SET'],
  ['ns', 'standard', 'percent', 0.10, 'Seremban/NS: SET x Price x 10% (部署前请再核对)'],
  ['johor', 'standard', 'fixed_per_set', 2.5, 'Johor: 暂按 RM2.50 x SET，部署前请与客户确认'],
  ['any', 'umbrella', 'fixed_per_set', 1, 'Umbrella: RM1 x SET，不分地区']
];

// ================= 初始化 =================

/**
 * 第一次使用前，在 Apps Script 编辑器里手动运行一次这个函数。
 * 会建好 4 个分页并灌入全年名单 / 默认价 / 抽佣规则。
 * 重复运行不会重复灌数据（已有数据就跳过）。
 */
function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_ORDERS, ORDERS_HEADERS);
  var configSheet = ensureSheet_(ss, SHEET_CONFIG, CONFIG_HEADERS);
  var ratesSheet = ensureSheet_(ss, SHEET_RATES, RATES_HEADERS);
  var pricesSheet = ensureSheet_(ss, SHEET_PRICES, PRICES_HEADERS);

  if (configSheet.getLastRow() < 2) {
    configSheet.getRange(2, 1, SEED_CONFIG.length, CONFIG_HEADERS.length).setValues(SEED_CONFIG);
  }
  if (ratesSheet.getLastRow() < 2) {
    ratesSheet.getRange(2, 1, SEED_RATES.length, RATES_HEADERS.length).setValues(SEED_RATES);
  }
  if (pricesSheet.getLastRow() < 2) {
    var priceRows = [];
    Object.keys(SEED_PRICES).forEach(function (branch) {
      Object.keys(SEED_PRICES[branch]).forEach(function (setType) {
        priceRows.push([branch, setType, SEED_PRICES[branch][setType]]);
      });
    });
    if (priceRows.length) {
      pricesSheet.getRange(2, 1, priceRows.length, PRICES_HEADERS.length).setValues(priceRows);
    }
  }

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('DRIVER_PHONE')) props.setProperty('DRIVER_PHONE', '');

  SpreadsheetApp.getUi().alert('初始化完成！已灌入 ' + SEED_CONFIG.length +
    ' 位 salesman。现在可以 部署 > 新增部署 > 网页应用。');
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ================= 网页入口 =================

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('订单 & 抽佣记录')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ================= 读取：启动数据 =================

function getBootstrapData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = sheetToObjects_(ss.getSheetByName(SHEET_CONFIG));
  var rates = sheetToObjects_(ss.getSheetByName(SHEET_RATES));
  var prices = sheetToObjects_(ss.getSheetByName(SHEET_PRICES));

  // salesmen 扁平列表：以名字为中心
  var salesmen = [];
  var branchesByRegion = {}; // region -> [branch,...]
  config.forEach(function (row) {
    if (!row.Salesman) {
      // 只有 branch 没 salesman 的行：也要让 branch 出现在下拉里
      if (row.Region && row.Branch) addBranchToMap_(branchesByRegion, row.Region, row.Branch);
      return;
    }
    salesmen.push({
      name: row.Salesman,
      branch: row.Branch,
      state: row.State,
      region: row.Region
    });
    addBranchToMap_(branchesByRegion, row.Region, row.Branch);
  });

  salesmen.sort(function (a, b) { return a.name < b.name ? -1 : a.name > b.name ? 1 : 0; });

  // prices: branch -> { setType: price }
  var priceMap = {};
  prices.forEach(function (p) {
    if (!p.Branch) return;
    if (!priceMap[p.Branch]) priceMap[p.Branch] = {};
    priceMap[p.Branch][String(p.SetType).toUpperCase()] = Number(p.Price) || 0;
  });

  var props = PropertiesService.getScriptProperties();

  return {
    salesmen: salesmen,
    branchesByRegion: branchesByRegion,
    prices: priceMap,
    rates: rates,
    regionLabels: REGION_LABELS,
    driverPhone: props.getProperty('DRIVER_PHONE') || '',
    userEmail: Session.getActiveUser().getEmail() || 'unknown'
  };
}

function addBranchToMap_(map, region, branch) {
  if (!map[region]) map[region] = [];
  if (branch && map[region].indexOf(branch) === -1) map[region].push(branch);
}

function sheetToObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return values.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// ================= 抽佣计算 =================

function calcAmounts_(region, item, price, setQty) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rates = sheetToObjects_(ss.getSheetByName(SHEET_RATES));
  price = Number(price) || 0;
  setQty = Number(setQty) || 0;
  var total = price * setQty;

  var isUmbrella = String(item).toLowerCase().indexOf('umbrella') !== -1;
  var rule = null;
  if (isUmbrella) {
    rule = rates.filter(function (r) { return String(r.Item).toLowerCase() === 'umbrella'; })[0];
  } else {
    rule = rates.filter(function (r) { return r.Region === region && r.Item === 'standard'; })[0];
  }

  var commission = 0;
  if (rule) {
    if (rule.Mode === 'percent') commission = total * Number(rule.Value);
    else if (rule.Mode === 'fixed_per_set') commission = setQty * Number(rule.Value);
  }
  return { total: round2_(total), commission: round2_(commission) };
}

function round2_(n) { return Math.round(n * 100) / 100; }

/** 前端实时预览，不写表 */
function previewAmounts(region, item, price, setQty) {
  return calcAmounts_(region, item, price, setQty);
}

// ================= 写入：订单 =================

function saveOrder(order) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureSheet_(ss, SHEET_ORDERS, ORDERS_HEADERS);

  var amounts = calcAmounts_(order.region, order.item, order.price, order.setQty);
  var lastRow = sheet.getLastRow();
  var newId = lastRow < 2 ? 1 : Number(sheet.getRange(lastRow, 1).getValue()) + 1;

  var date = order.date ? new Date(order.date) : new Date();
  var month = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');

  sheet.appendRow([
    newId, new Date(), date,
    order.region, order.state, order.branch, order.salesman,
    order.item, order.setDetail || '',
    Number(order.price) || 0, Number(order.setQty) || 0,
    amounts.total, amounts.commission,
    order.status || 'OP', order.paymentDate || '', month,
    Session.getActiveUser().getEmail() || 'unknown'
  ]);

  return { id: newId, total: amounts.total, commission: amounts.commission, month: month };
}

// ================= 写入：名单 / 价格 / 规则 =================

function addSalesman(region, state, branch, salesman) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureSheet_(ss, SHEET_CONFIG, CONFIG_HEADERS);
  var existing = sheetToObjects_(sheet);
  var dup = existing.some(function (r) {
    return r.Region === region && r.Branch === branch && r.Salesman === salesman;
  });
  if (!dup) sheet.appendRow([region, state, branch, salesman]);
  return getBootstrapData();
}

function addBranch(region, state, branch) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureSheet_(ss, SHEET_CONFIG, CONFIG_HEADERS);
  var existing = sheetToObjects_(sheet);
  var dup = existing.some(function (r) { return r.Region === region && r.Branch === branch; });
  if (!dup) sheet.appendRow([region, state, branch, '']);
  return getBootstrapData();
}

/** 新增或更新某 branch 某 set 类型的默认价 */
function setPrice(branch, setType, price) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureSheet_(ss, SHEET_PRICES, PRICES_HEADERS);
  setType = String(setType).toUpperCase();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === branch && String(data[i][1]).toUpperCase() === setType) {
      sheet.getRange(i + 1, 3).setValue(Number(price));
      return getBootstrapData();
    }
  }
  sheet.appendRow([branch, setType, Number(price)]);
  return getBootstrapData();
}

function updateRate(region, item, mode, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_RATES);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === region && data[i][1] === item) {
      sheet.getRange(i + 1, 3).setValue(mode);
      sheet.getRange(i + 1, 4).setValue(Number(value));
      return getBootstrapData();
    }
  }
  sheet.appendRow([region, item, mode, Number(value), '']);
  return getBootstrapData();
}

function setDriverPhone(phone) {
  PropertiesService.getScriptProperties().setProperty('DRIVER_PHONE', phone);
  return { ok: true };
}

// ================= 删除订单 =================

/** 按 ID 删除一笔订单。返回 {ok:true} 或 {ok:false} */
function deleteOrder(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return { ok: false };
  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (Number(ids[i][0]) === Number(id)) {
      sheet.deleteRow(i + 2);
      return { ok: true, id: id };
    }
  }
  return { ok: false };
}

// ================= 查找 / 汇总 / 仪表盘 =================

function searchOrders(filters) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = sheetToObjects_(ss.getSheetByName(SHEET_ORDERS));
  filters = filters || {};
  var kw = filters.keyword ? String(filters.keyword).toLowerCase() : '';

  return rows.filter(function (r) {
    if (filters.month && r.Month !== filters.month) return false;
    if (filters.region && r.Region !== filters.region) return false;
    if (filters.branch && r.Branch !== filters.branch) return false;
    if (filters.salesman && r.Salesman !== filters.salesman) return false;
    if (filters.status && r.Status !== filters.status) return false;
    if (kw && String(r.Salesman).toLowerCase().indexOf(kw) === -1 &&
      String(r.Branch).toLowerCase().indexOf(kw) === -1) return false;
    return true;
  }).map(mapOrderRow_).sort(function (a, b) { return b.id - a.id; });
}

function mapOrderRow_(r) {
  return {
    id: r.ID, date: formatDate_(r.Date), region: r.Region, state: r.State,
    branch: r.Branch, salesman: r.Salesman, item: r.Item, setDetail: r.SetDetail,
    price: r.Price, setQty: r.Set, total: r.Total, commission: r.Commission,
    status: r.Status, paymentDate: r.PaymentDate, month: r.Month
  };
}

function formatDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yy');
  }
  return d;
}

function getMonthlySummary(month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = sheetToObjects_(ss.getSheetByName(SHEET_ORDERS)).filter(function (r) {
    return r.Month === month;
  });

  var totalIncome = 0, totalCommission = 0, byBranch = {};
  rows.forEach(function (r) {
    totalIncome += Number(r.Total) || 0;
    totalCommission += Number(r.Commission) || 0;
    if (!byBranch[r.Branch]) byBranch[r.Branch] = { total: 0, commission: 0, count: 0 };
    byBranch[r.Branch].total += Number(r.Total) || 0;
    byBranch[r.Branch].commission += Number(r.Commission) || 0;
    byBranch[r.Branch].count += 1;
  });

  var branchList = Object.keys(byBranch).map(function (b) {
    return { branch: b, total: round2_(byBranch[b].total), commission: round2_(byBranch[b].commission), count: byBranch[b].count };
  }).sort(function (a, b) { return b.total - a.total; });

  return {
    month: month, orderCount: rows.length,
    totalIncome: round2_(totalIncome), totalCommission: round2_(totalCommission),
    byBranch: branchList
  };
}

/** 首页仪表盘：本月汇总 + 最近订单 */
function getDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = sheetToObjects_(ss.getSheetByName(SHEET_ORDERS));
  var thisMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

  var mTotal = 0, mComm = 0, mCount = 0, allTotal = 0, allComm = 0;
  rows.forEach(function (r) {
    allTotal += Number(r.Total) || 0;
    allComm += Number(r.Commission) || 0;
    if (r.Month === thisMonth) {
      mTotal += Number(r.Total) || 0;
      mComm += Number(r.Commission) || 0;
      mCount += 1;
    }
  });

  var recent = rows.map(mapOrderRow_).sort(function (a, b) { return b.id - a.id; }).slice(0, 8);

  return {
    month: thisMonth,
    monthTotal: round2_(mTotal), monthCommission: round2_(mComm), monthCount: mCount,
    allTotal: round2_(allTotal), allCommission: round2_(allComm), allCount: rows.length,
    recent: recent
  };
}

function listAvailableMonths() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rows = sheetToObjects_(ss.getSheetByName(SHEET_ORDERS));
  var months = {};
  rows.forEach(function (r) { if (r.Month) months[r.Month] = true; });
  var list = Object.keys(months).sort().reverse();
  if (list.length === 0) list = [Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM')];
  return list;
}
