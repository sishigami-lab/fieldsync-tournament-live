export type StaffRole="admin"|"referee";

const cookieName="fieldsync_staff";
const encoder=new TextEncoder();
const secret=()=>process.env.SUPABASE_SECRET_KEY||"";
const base64Url=(bytes:Uint8Array)=>btoa(String.fromCharCode(...bytes)).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/g,"");
const sign=async(value:string)=>{const key=await crypto.subtle.importKey("raw",encoder.encode(secret()),{name:"HMAC",hash:"SHA-256"},false,["sign"]);return base64Url(new Uint8Array(await crypto.subtle.sign("HMAC",key,encoder.encode(value))))};

export async function createStaffSession(role:StaffRole){
 const payload=base64Url(encoder.encode(JSON.stringify({role,exp:Date.now()+12*60*60*1000})));
 return `${payload}.${await sign(payload)}`;
}

export async function getStaffRole(request:Request):Promise<StaffRole|null>{
 const raw=request.headers.get("cookie")?.split(";").map(value=>value.trim()).find(value=>value.startsWith(`${cookieName}=`))?.slice(cookieName.length+1);
 if(!raw)return null;
 const [payload,signature]=raw.split(".");
 if(!payload||!signature||await sign(payload)!==signature)return null;
 try{
  const normalized=payload.replace(/-/g,"+").replace(/_/g,"/");
  const value=JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(normalized),character=>character.charCodeAt(0)))) as {role:StaffRole;exp:number};
  return value.exp>Date.now()&&(value.role==="admin"||value.role==="referee")?value.role:null;
 }catch{return null}
}

export const staffCookie=(token:string)=>`${cookieName}=${token}; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=43200`;
export const clearStaffCookie=()=>`${cookieName}=; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=0`;
export const forbidden=()=>Response.json({error:"この操作を行う権限がありません"},{status:403});
