import {sql} from "drizzle-orm";
import {index,text,sqliteTable} from "drizzle-orm/sqlite-core";

export const tournaments=sqliteTable("tournaments",{
 id:text("id").primaryKey(),
 name:text("name").notNull(),
 status:text("status").notNull().default("draft"),
 data:text("data").notNull(),
 createdAt:text("created_at").notNull().default(sql`CURRENT_TIMESTAMP`),
 updatedAt:text("updated_at").notNull().default(sql`CURRENT_TIMESTAMP`),
},table=>[index("idx_tournaments_updated_at").on(table.updatedAt)]);
