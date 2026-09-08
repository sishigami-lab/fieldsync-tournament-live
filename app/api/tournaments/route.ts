import {desc,eq} from "drizzle-orm";
import {getDb} from "../../../db";
import {tournaments} from "../../../db/schema";

const parse=(row:typeof tournaments.$inferSelect)=>({...row,data:JSON.parse(row.data)});
export async function GET(){const rows=await getDb().select().from(tournaments).orderBy(desc(tournaments.updatedAt));return Response.json({tournaments:rows.map(parse)})}
export async function POST(request:Request){const body=await request.json() as {name?:string;status?:string;data?:unknown},id=crypto.randomUUID(),now=new Date().toISOString();if(!body.name)return Response.json({error:"大会名が必要です"},{status:400});const [row]=await getDb().insert(tournaments).values({id,name:body.name,status:body.status||"draft",data:JSON.stringify(body.data||{}),createdAt:now,updatedAt:now}).returning();return Response.json({tournament:parse(row)},{status:201})}
export async function PUT(request:Request){const body=await request.json() as {id?:string;name?:string;status?:string;data?:unknown};if(!body.id||!body.name)return Response.json({error:"保存情報が不足しています"},{status:400});const [row]=await getDb().update(tournaments).set({name:body.name,status:body.status||"draft",data:JSON.stringify(body.data||{}),updatedAt:new Date().toISOString()}).where(eq(tournaments.id,body.id)).returning();return Response.json({tournament:parse(row)})}
export async function DELETE(request:Request){const id=new URL(request.url).searchParams.get("id");if(!id)return Response.json({error:"IDが必要です"},{status:400});await getDb().delete(tournaments).where(eq(tournaments.id,id));return Response.json({ok:true})}
