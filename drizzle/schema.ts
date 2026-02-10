import { int, json, mysqlEnum, mysqlTable, text, timestamp, varchar, decimal } from "drizzle-orm/mysql-core";

/**
 * Core user table backing auth flow.
 */
export const users = mysqlTable("users", {
  id: int("id").autoincrement().primaryKey(),
  openId: varchar("openId", { length: 64 }).notNull().unique(),
  name: text("name"),
  email: varchar("email", { length: 320 }),
  loginMethod: varchar("loginMethod", { length: 64 }),
  role: mysqlEnum("role", ["user", "admin"]).default("user").notNull(),
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
  lastSignedIn: timestamp("lastSignedIn").defaultNow().notNull(),
});

export type User = typeof users.$inferSelect;
export type InsertUser = typeof users.$inferInsert;

/**
 * Relatórios de visita técnica
 */
export const reports = mysqlTable("reports", {
  id: int("id").autoincrement().primaryKey(),
  userId: int("userId").notNull(),
  
  // Dados do Projetista
  projectistName: text("projectistName"),
  projectistCompany: text("projectistCompany"),
  projectistPhone: text("projectistPhone"),
  
  // Dados do Cliente
  clientCompanyName: text("clientCompanyName"),
  clientAddress: text("clientAddress"),
  clientCity: text("clientCity"),
  clientState: text("clientState"),
  clientZipCode: text("clientZipCode"),
  clientNeighborhood: text("clientNeighborhood"),
  clientSANumber: text("clientSANumber"),
  clientARTNumber: text("clientARTNumber"),
  
  // Informações da Central
  unitLocation: text("unitLocation"),
  isNewClient: mysqlEnum("isNewClient", ["yes", "no"]).default("yes"),
  isAsBuilt: mysqlEnum("isAsBuilt", ["yes", "no"]).default("no"),
  isAdequation: mysqlEnum("isAdequation", ["yes", "no"]).default("no"),
  
  // Checklist de Verificação
  checklist: json("checklist").$type<{
    centralReady: boolean | null;
    masonryReady: boolean | null;
    tankReady: boolean | null;
    networkExists: boolean | null;
    pointsQuantityChange: boolean | null;
    truckAccess: boolean | null;
    extintorValid: boolean | null;
    cellSignal: string | null;
    cellOperator: string | null;
    cellWifi: boolean | null;
  }>(),
  
  // Quantitativos/Escopo
  quantitatives: json("quantitatives").$type<{
    tankQuantity: string;
    networkBitola: string;
    networkMeterage: string;
    pointsQuantity: string;
    shelterQuantity: string;
    resistantWallArea: string;
    gateArea: string;
    slabArea: string;
  }>(),
  
  // Avaliação de Dificuldade
  logisticDifficulty: int("logisticDifficulty"),
  assemblyDifficulty: int("assemblyDifficulty"),
  
  // Observações e Ações
  observations: text("observations"),
  clientActions: text("clientActions"),
  supergasActions: text("supergasActions"),
  
  // Status do Relatório
  status: mysqlEnum("status", ["draft", "completed"]).default("draft"),
  
  // Timestamps
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
  completedAt: timestamp("completedAt"),
});

export type Report = typeof reports.$inferSelect;
export type InsertReport = typeof reports.$inferInsert;

/**
 * Fotos dos relatórios
 */
export const photos = mysqlTable("photos", {
  id: int("id").autoincrement().primaryKey(),
  reportId: int("reportId").notNull(),
  
  // URL da foto no S3
  photoUrl: text("photoUrl").notNull(),
  photoKey: text("photoKey").notNull(),
  
  // Descrição/Legenda da foto
  caption: text("caption"),
  
  // Ordem de exibição
  order: int("order").default(0),
  
  // Timestamps
  createdAt: timestamp("createdAt").defaultNow().notNull(),
});

export type Photo = typeof photos.$inferSelect;
export type InsertPhoto = typeof photos.$inferInsert;

/**
 * Rascunhos automáticos (salvamento automático)
 */
export const drafts = mysqlTable("drafts", {
  id: int("id").autoincrement().primaryKey(),
  reportId: int("reportId"),
  userId: int("userId").notNull(),
  
  // Dados do rascunho
  draftData: json("draftData").$type<Record<string, unknown>>(),
  
  // Timestamps
  createdAt: timestamp("createdAt").defaultNow().notNull(),
  updatedAt: timestamp("updatedAt").defaultNow().onUpdateNow().notNull(),
});

export type Draft = typeof drafts.$inferSelect;
export type InsertDraft = typeof drafts.$inferInsert;
