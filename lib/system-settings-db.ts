import {env} from "cloudflare:workers";
import {hashPin} from "./password";

export type SystemSettings={
 admin_password_salt:string;
 admin_password_hash:string;
 referee_password_salt:string;
 referee_password_hash:string;
 confirm_result:boolean;
 sound:boolean;
 vibration:boolean;
 live_notice:boolean;
};

type D1Result<T>={results?:T[]};
type D1Statement={bind:(...values:unknown[])=>D1Statement;first:<T>()=>Promise<T|null>;run:()=>Promise<unknown>};
type D1Database={prepare:(sql:string)=>D1Statement};

const database=()=>{
 const value=(env as unknown as {DB?:D1Database}).DB;
 if(!value)throw new Error("Cloudflare D1 binding DB is unavailable");
 return value;
};

const createTable=`CREATE TABLE IF NOT EXISTS system_settings (
 id TEXT PRIMARY KEY,
 admin_password_salt TEXT NOT NULL,
 admin_password_hash TEXT NOT NULL,
 referee_password_salt TEXT NOT NULL,
 referee_password_hash TEXT NOT NULL,
 confirm_result INTEGER NOT NULL DEFAULT 1,
 sound INTEGER NOT NULL DEFAULT 1,
 vibration INTEGER NOT NULL DEFAULT 1,
 live_notice INTEGER NOT NULL DEFAULT 1,
 updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
)`;

const normalize=(row:Record<string,unknown>):SystemSettings=>({
 admin_password_salt:String(row.admin_password_salt),
 admin_password_hash:String(row.admin_password_hash),
 referee_password_salt:String(row.referee_password_salt),
 referee_password_hash:String(row.referee_password_hash),
 confirm_result:Boolean(row.confirm_result),
 sound:Boolean(row.sound),
 vibration:Boolean(row.vibration),
 live_notice:Boolean(row.live_notice),
});

export async function getSystemSettings(){
 const db=database();
 await db.prepare(createTable).run();
 let row=await db.prepare("SELECT * FROM system_settings WHERE id = 'default'").first<Record<string,unknown>>();
 if(!row){
  const admin=await hashPin("1234"),referee=await hashPin("5678");
  await db.prepare("INSERT INTO system_settings (id, admin_password_salt, admin_password_hash, referee_password_salt, referee_password_hash) VALUES ('default', ?, ?, ?, ?)").bind(admin.salt,admin.hash,referee.salt,referee.hash).run();
  row=await db.prepare("SELECT * FROM system_settings WHERE id = 'default'").first<Record<string,unknown>>();
 }
 if(!row)throw new Error("System settings could not be initialized");
 return normalize(row);
}

export async function updateSystemSettings(patch:Partial<SystemSettings>){
 const entries=Object.entries(patch);
 if(!entries.length)return getSystemSettings();
 const columns=entries.map(([key])=>`${key} = ?`).join(", ");
 await database().prepare(`UPDATE system_settings SET ${columns}, updated_at = CURRENT_TIMESTAMP WHERE id = 'default'`).bind(...entries.map(([,value])=>typeof value==="boolean"?(value?1:0):value)).run();
 return getSystemSettings();
}
