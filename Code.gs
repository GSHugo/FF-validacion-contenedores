const SHEETS = {
  CONFIG: 'CONFIGURACION', CONTENEDORES: 'CONTENEDORES', EQUIPOS: 'EQUIPOS',
  TARIMAS: 'TARIMAS', CAPTURAS: 'CAPTURAS', REPORTE: 'REPORTE_GENERAL'
};

const HEADERS = {
  CONFIGURACION: ['CLAVE','VALOR','DESCRIPCION'],
  CONTENEDORES: ['ID_CONTENEDOR','REFERENCIA_CONTENEDOR','ORDEN_COMPRA','ESTADO','FECHA_APERTURA','FECHA_CIERRE','CARPETA_DRIVE_ID'],
  EQUIPOS: ['ID_EQUIPO','ID_CONTENEDOR','NOMBRE_EQUIPO','ESTADO','CREADO_POR','FECHA_CREACION'],
  TARIMAS: ['ID_TARIMA','ID_CONTENEDOR','ID_EQUIPO','CODIGO_TARIMA','ESTADO','CREADO_POR','FECHA_CREACION','FECHA_FINALIZACION'],
  CAPTURAS: ['ID_CAPTURA','ID_CONTENEDOR','ID_EQUIPO','ID_TARIMA','CODIGO_TARIMA','NUMERO_CAJA','VALIDADOR','REFERENCIA_INTERNA','EAN','DESCRIPCION','PIEZAS_CAJA','COLOR','TALLA','ALTO_CM','LARGO_CM','ANCHO_CM','VOLUMEN_CM3','FOTO_ITEM_URL','FOTO_ITEM_ID','FOTO_MASTER_URL','FOTO_MASTER_ID','OBSERVACIONES','FECHA_CREACION','FECHA_MODIFICACION','MODIFICADO_POR','ESTADO'],
  REPORTE_GENERAL: ['ID_CONTENEDOR','REFERENCIA_CONTENEDOR','CODIGO_TARIMA','SKU_ODO','IMAGEN','REFERENCIA_INTERNA','EAN','DESCRIPCION','CAJAS_TOTALES','PIEZAS_POR_CAJA','PIEZAS_TOTALES','COLOR','TALLA','ALTO_CM','LARGO_CM','ANCHO_CM','VOLUMEN_CM3','VALIDADORES','FOTO_ITEM_URL','FOTO_MASTER_URL']
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Validación de Contenedores')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(name => ensureSheet_(ss, name, HEADERS[name]));
  const config = ss.getSheetByName(SHEETS.CONFIG);
  if (config.getLastRow() === 1) {
    const root = DriveApp.createFolder('VALIDACIONES_CONTENEDORES');
    config.getRange(2,1,3,3).setValues([
      ['ROOT_FOLDER_ID', root.getId(), 'Carpeta raíz para fotografías'],
      ['APP_VERSION', '1.0.0', 'Versión instalada'],
      ['POLL_SECONDS', '4', 'Actualización de tarima compartida']
    ]);
  }
  styleSheets_();
  return {ok:true, spreadsheetId:ss.getId(), message:'Sistema configurado'};
}

function getBootstrap() {
  assertSetup_();
  return {containers: rows_(SHEETS.CONTENEDORES).filter(r => r.ESTADO !== 'FINALIZADO')};
}

function createContainer(data) {
  return withLock_(() => {
    assertText_(data.reference, 'Referencia del contenedor');
    const existing = rows_(SHEETS.CONTENEDORES).find(r => norm_(r.REFERENCIA_CONTENEDOR) === norm_(data.reference) && r.ESTADO !== 'FINALIZADO');
    if (existing) return existing;
    const id = id_('CONT');
    const root = DriveApp.getFolderById(config_('ROOT_FOLDER_ID'));
    const folder = root.createFolder(safeName_(data.reference));
    folder.createFolder('FOTOS_ITEMS');
    folder.createFolder('FOTOS_CAJAS_MASTER');
    append_(SHEETS.CONTENEDORES, [id,data.reference,String(data.purchaseOrder||''),'EN PROCESO',new Date(),' ',folder.getId()]);
    return rows_(SHEETS.CONTENEDORES).find(r => r.ID_CONTENEDOR === id);
  });
}

function openTeam(data) {
  return withLock_(() => {
    assertText_(data.containerId, 'Contenedor'); assertText_(data.teamName, 'Nombre del equipo'); assertText_(data.person, 'Nombre de la persona');
    const container = find_(SHEETS.CONTENEDORES,'ID_CONTENEDOR',data.containerId);
    if (!container || container.ESTADO === 'FINALIZADO') throw new Error('El contenedor no está disponible.');
    let team = rows_(SHEETS.EQUIPOS).find(r => r.ID_CONTENEDOR === data.containerId && norm_(r.NOMBRE_EQUIPO) === norm_(data.teamName) && r.ESTADO === 'ACTIVO');
    if (!team) {
      const teamId = id_('EQ');
      append_(SHEETS.EQUIPOS,[teamId,data.containerId,data.teamName,'ACTIVO',data.person,new Date()]);
      team = find_(SHEETS.EQUIPOS,'ID_EQUIPO',teamId);
    }
    return {team, pallets: listPallets({containerId:data.containerId,teamId:team.ID_EQUIPO})};
  });
}

function createPallet(data) {
  return withLock_(() => {
    assertTeam_(data.containerId,data.teamId);
    const prefix = safeName_(find_(SHEETS.CONTENEDORES,'ID_CONTENEDOR',data.containerId).REFERENCIA_CONTENEDOR).slice(0,10);
    let code;
    do { code = prefix + '-T-' + Math.random().toString(36).slice(2,6).toUpperCase(); }
    while (rows_(SHEETS.TARIMAS).some(r => r.CODIGO_TARIMA === code));
    const id = id_('TAR');
    append_(SHEETS.TARIMAS,[id,data.containerId,data.teamId,code,'EN PROCESO',data.person,new Date(),'']);
    return find_(SHEETS.TARIMAS,'ID_TARIMA',id);
  });
}

function listPallets(data) {
  assertTeam_(data.containerId,data.teamId);
  return rows_(SHEETS.TARIMAS).filter(r => r.ID_CONTENEDOR === data.containerId && r.ID_EQUIPO === data.teamId);
}

function getPallet(data) {
  const pallet = assertPallet_(data.palletId,data.teamId);
  const captures = rows_(SHEETS.CAPTURAS).filter(r => r.ID_TARIMA === data.palletId && r.ESTADO !== 'ELIMINADO');
  return {pallet, captures, totals:{boxes:captures.length,pieces:captures.reduce((n,r)=>n+Number(r.PIEZAS_CAJA||0),0)}};
}

function saveCapture(data) {
  return withLock_(() => {
    const pallet = assertPallet_(data.palletId,data.teamId);
    if (pallet.ESTADO !== 'EN PROCESO') throw new Error('La tarima ya fue finalizada y está bloqueada.');
    if (!String(data.internalRef||'').trim() && !String(data.ean||'').trim()) throw new Error('Captura la referencia interna, el EAN o ambos.');
    if (Number(data.pieces) < 1) throw new Error('Las piezas de la caja deben ser mayores a cero.');
    assertText_(data.description,'Descripción');
    const existing = data.captureId ? find_(SHEETS.CAPTURAS,'ID_CAPTURA',data.captureId) : null;
    if (existing && existing.ID_EQUIPO !== data.teamId) throw new Error('El registro pertenece a otro equipo.');
    const container = find_(SHEETS.CONTENEDORES,'ID_CONTENEDOR',pallet.ID_CONTENEDOR);
    const photos = savePhotos_(container, data, existing);
    const dims = [data.height,data.length,data.width].map(v => Number(v||0));
    const captureId = existing ? existing.ID_CAPTURA : id_('CAP');
    const boxNo = existing ? existing.NUMERO_CAJA : rows_(SHEETS.CAPTURAS).filter(r=>r.ID_TARIMA===pallet.ID_TARIMA && r.ESTADO!=='ELIMINADO').length+1;
    const row = [captureId,pallet.ID_CONTENEDOR,data.teamId,pallet.ID_TARIMA,pallet.CODIGO_TARIMA,boxNo,data.person,String(data.internalRef||''),String(data.ean||''),data.description,Number(data.pieces),String(data.color||''),String(data.size||''),dims[0],dims[1],dims[2],dims[0]*dims[1]*dims[2],photos.item.url,photos.item.id,photos.master.url,photos.master.id,String(data.notes||''),existing?existing.FECHA_CREACION:new Date(),new Date(),data.person,'ACTIVO'];
    existing ? replaceById_(SHEETS.CAPTURAS,'ID_CAPTURA',captureId,row) : append_(SHEETS.CAPTURAS,row);
    rebuildReport_();
    return getPallet({palletId:pallet.ID_TARIMA,teamId:data.teamId});
  });
}

function finalizePallet(data) {
  return withLock_(() => {
    const pallet = assertPallet_(data.palletId,data.teamId);
    const captures = rows_(SHEETS.CAPTURAS).filter(r=>r.ID_TARIMA===data.palletId && r.ESTADO!=='ELIMINADO');
    if (!captures.length) throw new Error('No se puede finalizar una tarima vacía.');
    updateFields_(SHEETS.TARIMAS,'ID_TARIMA',data.palletId,{ESTADO:'FINALIZADA',FECHA_FINALIZACION:new Date()});
    rebuildReport_();
    return getPallet({palletId:data.palletId,teamId:data.teamId});
  });
}

function savePhotos_(container,data,existing) {
  if (!data.itemPhoto && !existing) throw new Error('La foto del artículo es obligatoria.');
  if (!data.masterPhoto && !existing) throw new Error('La foto de la caja máster es obligatoria.');
  const root = DriveApp.getFolderById(container.CARPETA_DRIVE_ID);
  const itemFolder = childFolder_(root,'FOTOS_ITEMS');
  const masterFolder = childFolder_(root,'FOTOS_CAJAS_MASTER');
  return {
    item: data.itemPhoto ? saveDataUrl_(itemFolder,data.itemPhoto,'ITEM_'+id_('IMG')) : {id:existing.FOTO_ITEM_ID,url:existing.FOTO_ITEM_URL},
    master: data.masterPhoto ? saveDataUrl_(masterFolder,data.masterPhoto,'MASTER_'+id_('IMG')) : {id:existing.FOTO_MASTER_ID,url:existing.FOTO_MASTER_URL}
  };
}

function saveDataUrl_(folder,dataUrl,name) {
  const m = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) throw new Error('Formato de fotografía no válido.');
  const ext = m[1].includes('png') ? 'png' : 'jpg';
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]),m[1],name+'.'+ext));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW); } catch(e) {}
  return {id:file.getId(),url:'https://drive.google.com/uc?export=view&id='+file.getId()};
}

function rebuildReport_() {
  const captures = rows_(SHEETS.CAPTURAS).filter(r=>r.ESTADO!=='ELIMINADO');
  const containers = Object.fromEntries(rows_(SHEETS.CONTENEDORES).map(r=>[r.ID_CONTENEDOR,r]));
  const groups = {};
  captures.forEach(r => {
    const key=[r.ID_CONTENEDOR,r.CODIGO_TARIMA,norm_(r.REFERENCIA_INTERNA),norm_(r.EAN),norm_(r.DESCRIPCION),norm_(r.COLOR),norm_(r.TALLA),r.ALTO_CM,r.LARGO_CM,r.ANCHO_CM].join('|');
    if(!groups[key]) groups[key]={...r,boxes:0,pieces:0,pcs:[],people:new Set(),masters:[]};
    const g=groups[key]; g.boxes++; g.pieces+=Number(r.PIEZAS_CAJA||0); g.pcs.push(Number(r.PIEZAS_CAJA||0)); g.people.add(r.VALIDADOR); g.masters.push(r.FOTO_MASTER_URL);
  });
  const rows=Object.values(groups).map(g=>[g.ID_CONTENEDOR,containers[g.ID_CONTENEDOR]?.REFERENCIA_CONTENEDOR||'',g.CODIGO_TARIMA,'','=IMAGE("'+g.FOTO_ITEM_URL+'")',g.REFERENCIA_INTERNA,g.EAN,g.DESCRIPCION,g.boxes,[...new Set(g.pcs)].join(', '),g.pieces,g.COLOR,g.TALLA,g.ALTO_CM,g.LARGO_CM,g.ANCHO_CM,g.VOLUMEN_CM3,[...g.people].join(', '),g.FOTO_ITEM_URL,[...new Set(g.masters)].join('\n')]);
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.REPORTE); sh.getRange(2,1,Math.max(sh.getMaxRows()-1,1),sh.getLastColumn()).clearContent();
  if(rows.length) sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  Object.values(containers).forEach(c => syncContainerView_(c, rows.filter(r => r[0] === c.ID_CONTENEDOR)));
}

function adminRefreshReport() { rebuildReport_(); return 'Reporte y vistas actualizados'; }
function syncContainerView_(container, data) {
  const ss=SpreadsheetApp.getActive();
  const name=('VISTA_'+safeName_(container.REFERENCIA_CONTENEDOR)).slice(0,99);
  let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
  const headers=HEADERS.REPORTE_GENERAL;
  sh.clear(); sh.getRange(1,1,1,headers.length).setValues([headers]);
  if(data.length) sh.getRange(2,1,data.length,data[0].length).setValues(data);
  sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setBackground('#16324F').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
  sh.autoResizeColumns(1,headers.length); if(!sh.getFilter() && sh.getLastRow()>1) sh.getDataRange().createFilter();
}

function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);return sh;}
function styleSheets_(){const ss=SpreadsheetApp.getActive();Object.keys(HEADERS).forEach(n=>{const sh=ss.getSheetByName(n);sh.setFrozenRows(1);sh.getRange(1,1,1,HEADERS[n].length).setBackground('#16324F').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);if(!sh.getFilter()&&sh.getLastRow()>1)sh.getDataRange().createFilter();sh.autoResizeColumns(1,HEADERS[n].length);});}
function rows_(name){const sh=SpreadsheetApp.getActive().getSheetByName(name);const vals=sh.getDataRange().getValues();if(vals.length<2)return[];return vals.slice(1).filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>Object.fromEntries(vals[0].map((h,i)=>[h,r[i]])));}
function append_(name,row){SpreadsheetApp.getActive().getSheetByName(name).appendRow(row);}
function find_(name,key,value){return rows_(name).find(r=>String(r[key])===String(value));}
function replaceById_(name,key,value,row){const sh=SpreadsheetApp.getActive().getSheetByName(name),vals=sh.getDataRange().getValues(),idx=vals[0].indexOf(key);for(let i=1;i<vals.length;i++)if(String(vals[i][idx])===String(value)){sh.getRange(i+1,1,1,row.length).setValues([row]);return;}throw new Error('Registro no encontrado.');}
function updateFields_(name,key,value,changes){const sh=SpreadsheetApp.getActive().getSheetByName(name),vals=sh.getDataRange().getValues(),idx=vals[0].indexOf(key);for(let i=1;i<vals.length;i++)if(String(vals[i][idx])===String(value)){Object.entries(changes).forEach(([k,v])=>{const c=vals[0].indexOf(k);if(c>=0)sh.getRange(i+1,c+1).setValue(v);});return;}throw new Error('Registro no encontrado.');}
function assertTeam_(containerId,teamId){const t=find_(SHEETS.EQUIPOS,'ID_EQUIPO',teamId);if(!t||t.ID_CONTENEDOR!==containerId||t.ESTADO!=='ACTIVO')throw new Error('Equipo no válido para este contenedor.');return t;}
function assertPallet_(palletId,teamId){const p=find_(SHEETS.TARIMAS,'ID_TARIMA',palletId);if(!p||p.ID_EQUIPO!==teamId)throw new Error('La tarima no pertenece a este equipo.');return p;}
function assertText_(v,n){if(!String(v||'').trim())throw new Error(n+' es obligatorio.');}
function assertSetup_(){if(!SpreadsheetApp.getActive().getSheetByName(SHEETS.CONFIG))throw new Error('Ejecuta setupSystem() una vez desde Apps Script.');}
function config_(key){const r=rows_(SHEETS.CONFIG).find(x=>x.CLAVE===key);if(!r)throw new Error('Falta configuración: '+key);return r.VALOR;}
function childFolder_(parent,name){const it=parent.getFoldersByName(name);return it.hasNext()?it.next():parent.createFolder(name);}
function id_(p){return p+'-'+Utilities.getUuid().slice(0,8).toUpperCase();}
function norm_(v){return String(v||'').trim().toUpperCase();}
function safeName_(v){return String(v||'CONTENEDOR').trim().replace(/[^a-zA-Z0-9_-]+/g,'_');}
function withLock_(fn){const lock=LockService.getScriptLock();lock.waitLock(30000);try{return fn();}finally{lock.releaseLock();}}
