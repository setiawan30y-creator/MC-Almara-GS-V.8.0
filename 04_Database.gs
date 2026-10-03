const MC_Database={ensureSheet:function(n,h){var ss=this.ss(),s=ss.getSheetByName(n);if(!s){s=ss.insertSheet(n);s.getRange(1,1,1,h.length).setValues([h]);s.setFrozenRows(1)}return s},ss(){const id=PropertiesService.getScriptProperties().getProperty(MC_CONFIG.DB_PROPERTY);MC_Utils.require(id,'Database belum di-setup');return SpreadsheetApp.openById(id)},sheet(n){const s=this.ss().getSheetByName(n);MC_Utils.require(s,'Sheet '+n+' tidak ditemukan');return s},rows(n){const v=this.sheet(n).getDataRange().getValues();if(v.length<2)return[];return v.slice(1).filter(r=>r.some(x=>x!==''&&x!==null)).map(r=>Object.fromEntries(v[0].map((k,i)=>[k,r[i]])))},find(n,k,v){return this.rows(n).find(x=>String(x[k])===String(v))||null},update(n,k,v,o){const s=this.sheet(n),data=s.getDataRange().getValues(),h=data[0],idx=h.indexOf(k);MC_Utils.require(idx>=0,'Kolom '+k+' tidak ditemukan');for(let r=1;r<data.length;r++){if(String(data[r][idx])===String(v)){s.getRange(r+1,1,1,h.length).setValues([h.map(function(key,i){return o[key]??data[r][i]})]);return o}}throw new Error('Record tidak ditemukan')},insert(n,o){const s=this.sheet(n),h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];s.appendRow(h.map(k=>o[k]??''));return o},removeWhere(n,k,v){const s=this.sheet(n),data=s.getDataRange().getValues(),h=data[0],idx=h.indexOf(k);MC_Utils.require(idx>=0,'Kolom '+k+' tidak ditemukan');for(let r=data.length-1;r>=1;r--){if(String(data[r][idx])===String(v))s.deleteRow(r+1)}return true},audit(m,a,id,o,n,r){this.insert('AUDIT_LOG',{ID:Utilities.getUuid(),USER_ID:Session.getActiveUser().getEmail()||'system',TIMESTAMP:new Date(),MODULE:m,ACTION:a,RECORD_ID:id,OLD_VALUE:JSON.stringify(o??null),NEW_VALUE:JSON.stringify(n??null),IP:'',SESSION_ID:Session.getTemporaryActiveUserKey(),REASON:r||''})},install(){
  var configuredId=PropertiesService.getScriptProperties().getProperty(MC_CONFIG.DB_PROPERTY);
  var ss=SpreadsheetApp.getActiveSpreadsheet();

  if(!ss && configuredId){
    try{ss=SpreadsheetApp.openById(configuredId)}catch(e){}
  }

  MC_Utils.require(ss,'Spreadsheet database tidak ditemukan');

  PropertiesService.getScriptProperties().setProperty(MC_CONFIG.DB_PROPERTY,ss.getId());

  var names=Object.keys(MC_SHEETS);
  names.forEach(function(n){
    var s=ss.getSheetByName(n);
    if(!s)s=ss.insertSheet(n);
    var h=MC_SHEETS[n];

    if(s.getLastColumn()===0){
      s.getRange(1,1,1,h.length).setValues([h]);
    }else{
      var current=s.getRange(1,1,1,Math.max(s.getLastColumn(),h.length)).getValues()[0];
      var missing=current.slice(0,h.length).some(function(v,i){return String(v||'')!==String(h[i])});
      if(missing && s.getLastRow()<=1){
        s.clear();
        s.getRange(1,1,1,h.length).setValues([h]);
      }
    }
    s.setFrozenRows(1);
  });

  this.seed();

  SpreadsheetApp.flush();

  return{
    ok:true,
    spreadsheetId:ss.getId(),
    url:ss.getUrl(),
    sheetCount:names.length
  };
},
seed(){
  var ss=this.ss();

  function readRows(name){
    var s=ss.getSheetByName(name);
    var data=s.getDataRange().getValues();
    if(data.length<2)return[];
    return data.slice(1).filter(function(r){
      return r.some(function(x){return x!==''&&x!==null});
    }).map(function(r){
      return Object.fromEntries(data[0].map(function(k,i){return[k,r[i]]}));
    });
  }

  function appendBatch(name,objects){
    if(!objects.length)return;
    var s=ss.getSheetByName(name);
    var h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
    var values=objects.map(function(o){
      return h.map(function(k){return o[k]??''});
    });
    s.getRange(s.getLastRow()+1,1,values.length,h.length).setValues(values);
  }

  var roles=readRows('ROLES');
  if(!roles.length){
    appendBatch('ROLES',MC_CONFIG.ROLES.map(function(n,i){
      return{ID:'ROLE-'+(i+1),NAME:n,DESCRIPTION:'',STATUS:'ACTIVE'};
    }));
  }

  var roleRows=readRows('ROLES');
  var roleIds={};
  roleRows.forEach(function(r){roleIds[r.NAME]=r.ID});

  var perms=readRows('PERMISSIONS');
  var existingPerm={};
  perms.forEach(function(x){
    existingPerm[String(x.ROLE_ID)+'|'+String(x.MODULE)+'|'+String(x.PERMISSION)]=true;
  });

  var matrix={
    Admin:['view','create','edit','delete','print','export','approve','closing','rate_management','transaction_edit','wa'],
    Manager:['view','create','edit','print','export','approve','closing','rate_management','transaction_edit','wa'],
    Teller:['view','create','edit','print','transaction_edit','wa'],
    Kasir:['view','create','edit','print','wa'],
    Finance:['view','create','edit','print','export','approve','closing'],
    Auditor:['view','print','export']
  };

  var newPerms=[];
  MC_CONFIG.ROLES.forEach(function(name){
    (matrix[name]||['view']).forEach(function(permission){
      MC_CONFIG.MODULES.forEach(function(module){
        var roleId=roleIds[name];
        var key=String(roleId)+'|'+module+'|'+permission;
        if(!existingPerm[key]){
          newPerms.push({
            ID:Utilities.getUuid(),
            ROLE_ID:roleId,
            PERMISSION:permission,
            MODULE:module,
            STATUS:'ACTIVE'
          });
          existingPerm[key]=true;
        }
      });
    });
  });
  appendBatch('PERMISSIONS',newPerms);

  var configs=readRows('CONFIG');
  if(!configs.some(function(x){return x.KEY==='TRANSACTION_THRESHOLD_USD'})){
    appendBatch('CONFIG',[{
      KEY:'TRANSACTION_THRESHOLD_USD',
      VALUE:'10000',
      DESCRIPTION:'Monthly customer transaction threshold in USD equivalent'
    }]);
  }

  var currencies=readRows('CURRENCY_MASTER');
  if(!currencies.length){
    appendBatch('CURRENCY_MASTER',[
      ['USD','US Dollar',2],
      ['EUR','Euro',2],
      ['JPY','Japanese Yen',0],
      ['GBP','Pound Sterling',2],
      ['SGD','Singapore Dollar',2],
      ['AUD','Australian Dollar',2],
      ['CNY','Chinese Yuan',2]
    ].map(function(x,i){
      return{
        ID:'CUR-'+(i+1),
        ISO_CODE:x[0],
        NAME:x[1],
        COUNTRY:'',
        FLAG:'',
        SYMBOL:'',
        DECIMAL_PRECISION:x[2],
        STATUS:'ACTIVE',
        DISPLAY_ORDER:i+1
      };
    }));
  }
}};