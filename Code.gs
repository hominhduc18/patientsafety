// Dán đè vào Apps Script, rồi Triển khai > Quản lý bản triển khai > Sửa > Phiên bản mới > Triển khai (URL giữ nguyên)
function sh(n){const s=SpreadsheetApp.getActive();return s.getSheetByName(n)||s.insertSheet(n)}
function rows(n){return sh(n).getDataRange().getValues().filter(r=>r[0]!=='').map(r=>JSON.parse(r[1]))}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function upsert(n,o,k){const s=sh(n),v=s.getDataRange().getValues();for(let i=0;i<v.length;i++)if(v[i][0]==o[k]){s.getRange(i+1,2).setValue(JSON.stringify(o));return}s.appendRow([o[k],JSON.stringify(o)])}
function del(n,id){const s=sh(n),v=s.getDataRange().getValues();for(let i=v.length-1;i>=0;i--)if(v[i][0]==id)s.deleteRow(i+1)}
function sha(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('')}
function auth(u,p){const us=rows('Users');if(!us.length){const a={u:'admin',name:'Quản trị',role:'admin',p:sha('12345')};upsert('Users',a,'u');us.push(a)}return us.find(x=>x.u==u&&x.p==p)}
function doGet(){return out({ok:1})}
function doPost(e){const l=LockService.getScriptLock();l.waitLock(20000);try{
  const d=JSON.parse(e.postData.contents),me=auth(d.u,d.p);if(!me)return out({err:'Sai tên đăng nhập hoặc mật khẩu'});
  const ad=me.role=='admin',a=d.action;
  if(a=='save_rec')upsert('Records',d.rec,'id');
  else if(a=='save_form'&&ad)upsert('Forms',d.form,'id');
  else if(a=='del_form'&&ad)del('Forms',d.id);
  else if(a=='del_rec'&&ad)del('Records',d.id);
  else if(a=='save_user'&&ad)upsert('Users',d.user,'u');
  else if(a=='del_user'&&ad&&d.id!=me.u)del('Users',d.id);
  else if(a=='change_pass'){me.p=d.np;upsert('Users',me,'u')}
  else if(a!='login'&&a!='load')return out({err:'Không đủ quyền'});
  return out({me:{u:me.u,name:me.name,role:me.role},forms:rows('Forms'),recs:rows('Records'),users:ad?rows('Users').map(x=>({u:x.u,name:x.name,role:x.role})):[]})
}finally{l.releaseLock()}}
