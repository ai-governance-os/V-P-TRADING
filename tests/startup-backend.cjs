const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
let connections=0;const reads={};
const tables={USERS:[['NAME','PIN','ROLE','ACTIVE'],['TEST','2468','admin','YES'],['DRIVER','3579','driver','YES']],DRIVER:[['DRIVER_NAME','PHONE','ALLOWANCE_PER_MONTH'],['TEST','012-345',50]],CONFIG:[['KEY','VALUE'],['COMPANY','TEST']],SET_PRICE:[['SET_TYPE','UNIT_PRICE'],['BAG',10]],SALESMAN:[['SALESMAN','BRANCH'],['TEST','BRANCH']],BRANCH:[['BRANCH','BRAND','ACTIVE'],['TEST','PROTON','YES'],['OLD','HONDA','NO']]};
const spreadsheet={getSheetByName(name){
  return {getDataRange(){return {getValues(){
    reads[name]=(reads[name]||0)+1;return tables[name]||[];
  }}}};
}};
const c={
  SpreadsheetApp:{getActiveSpreadsheet(){connections++;return spreadsheet}},
  Session:{getScriptTimeZone:()=> 'Asia/Kuala_Lumpur'},
  PropertiesService:{getScriptProperties:()=>({getProperty:()=> '1'})},
  CacheService:{getScriptCache:()=>({get:()=>null,put(){}})}
};
vm.createContext(c);vm.runInContext(fs.readFileSync('gas/Code.js','utf8'),c);
assert.equal(connections,0,'loading page code does not connect to Sheets');
let denied=c.getAdminSettings('3579');
assert.equal(denied.ok,false,'non-admin cannot read account data');
assert.equal(reads.DRIVER||0,0,'denied settings does not read driver settings');
reads.USERS=0;
const settings=c.getAdminSettings('2468');
assert.equal(settings.ok,true);assert.equal(reads.USERS,1,'settings authenticates once');
assert.equal(reads.DRIVER,1);assert.equal(settings.driver.phone,'012345');
assert.equal(settings.users.length,2);
reads.BRANCH=0;
const boot=c.bootstrap();
assert.equal(reads.BRANCH,1,'bootstrap reuses branch data for brands');
assert.equal(boot.branches.length,1,'inactive branches stay excluded from choices');
assert.ok(boot.brands.includes('HONDA'),'historical brands stay included');
assert.equal(connections,1,'all reads reuse one connection');
console.log('12 startup/settings backend assertions passed');
