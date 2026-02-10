import { eq, and, desc } from "drizzle-orm";
import { drizzle } from "drizzle-orm/mysql2";
import { InsertUser, users, Report, InsertReport, reports, Photo, InsertPhoto, photos, Draft, InsertDraft, drafts } from "../drizzle/schema";
import { ENV } from './_core/env';

let _db: ReturnType<typeof drizzle> | null = null;

// Lazily create the drizzle instance so local tooling can run without a DB.
export async function getDb() {
  if (!_db && process.env.DATABASE_URL) {
    try {
      _db = drizzle(process.env.DATABASE_URL);
    } catch (error) {
      console.warn("[Database] Failed to connect:", error);
      _db = null;
    }
  }
  return _db;
}

export async function upsertUser(user: InsertUser): Promise<void> {
  if (!user.openId) {
    throw new Error("User openId is required for upsert");
  }

  const db = await getDb();
  if (!db) {
    console.warn("[Database] Cannot upsert user: database not available");
    return;
  }

  try {
    const values: InsertUser = {
      openId: user.openId,
    };
    const updateSet: Record<string, unknown> = {};

    const textFields = ["name", "email", "loginMethod"] as const;
    type TextField = (typeof textFields)[number];

    const assignNullable = (field: TextField) => {
      const value = user[field];
      if (value === undefined) return;
      const normalized = value ?? null;
      values[field] = normalized;
      updateSet[field] = normalized;
    };

    textFields.forEach(assignNullable);

    if (user.lastSignedIn !== undefined) {
      values.lastSignedIn = user.lastSignedIn;
      updateSet.lastSignedIn = user.lastSignedIn;
    }
    if (user.role !== undefined) {
      values.role = user.role;
      updateSet.role = user.role;
    } else if (user.openId === ENV.ownerOpenId) {
      values.role = 'admin';
      updateSet.role = 'admin';
    }

    if (!values.lastSignedIn) {
      values.lastSignedIn = new Date();
    }

    if (Object.keys(updateSet).length === 0) {
      updateSet.lastSignedIn = new Date();
    }

    await db.insert(users).values(values).onDuplicateKeyUpdate({
      set: updateSet,
    });
  } catch (error) {
    console.error("[Database] Failed to upsert user:", error);
    throw error;
  }
}

export async function getUserByOpenId(openId: string) {
  const db = await getDb();
  if (!db) {
    console.warn("[Database] Cannot get user: database not available");
    return undefined;
  }

  const result = await db.select().from(users).where(eq(users.openId, openId)).limit(1);

  return result.length > 0 ? result[0] : undefined;
}

// ============ REPORTS ============

export async function createReport(userId: number, data: Omit<InsertReport, 'userId'>) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db.insert(reports).values({
    ...data,
    userId,
  });

  return result;
}

export async function getReportById(reportId: number) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db.select().from(reports).where(eq(reports.id, reportId)).limit(1);
  return result.length > 0 ? result[0] : null;
}

export async function getReportsByUserId(userId: number, limit: number = 50, offset: number = 0) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db
    .select()
    .from(reports)
    .where(eq(reports.userId, userId))
    .orderBy(desc(reports.createdAt))
    .limit(limit)
    .offset(offset);

  return result;
}

export async function updateReport(reportId: number, data: Partial<InsertReport>) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db
    .update(reports)
    .set({
      ...data,
      updatedAt: new Date(),
    })
    .where(eq(reports.id, reportId));

  return result;
}

export async function deleteReport(reportId: number) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  // Delete photos first
  await db.delete(photos).where(eq(photos.reportId, reportId));

  // Delete drafts
  await db.delete(drafts).where(eq(drafts.reportId, reportId));

  // Delete report
  const result = await db.delete(reports).where(eq(reports.id, reportId));

  return result;
}

// ============ PHOTOS ============

export async function addPhoto(reportId: number, data: Omit<InsertPhoto, 'reportId'>) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db.insert(photos).values({
    ...data,
    reportId,
  });

  return result;
}

export async function getPhotosByReportId(reportId: number) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db
    .select()
    .from(photos)
    .where(eq(photos.reportId, reportId))
    .orderBy(photos.order);

  return result;
}

export async function updatePhoto(photoId: number, data: Partial<InsertPhoto>) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db
    .update(photos)
    .set(data)
    .where(eq(photos.id, photoId));

  return result;
}

export async function deletePhoto(photoId: number) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db.delete(photos).where(eq(photos.id, photoId));

  return result;
}

// ============ DRAFTS ============

export async function saveDraft(userId: number, reportId: number | null, draftData: Record<string, unknown>) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  // Check if draft exists
  const existingDraft = reportId
    ? await db.select().from(drafts).where(eq(drafts.reportId, reportId)).limit(1)
    : null;

  if (existingDraft && existingDraft.length > 0) {
    // Update existing draft
    return await db
      .update(drafts)
      .set({
        draftData,
        updatedAt: new Date(),
      })
      .where(eq(drafts.id, existingDraft[0].id));
  } else {
    // Create new draft
    return await db.insert(drafts).values({
      userId,
      reportId: reportId || undefined,
      draftData,
    });
  }
}

export async function getDraftByReportId(reportId: number) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db.select().from(drafts).where(eq(drafts.reportId, reportId)).limit(1);
  return result.length > 0 ? result[0] : null;
}

export async function getDraftsByUserId(userId: number) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db
    .select()
    .from(drafts)
    .where(eq(drafts.userId, userId))
    .orderBy(desc(drafts.updatedAt));

  return result;
}

export async function deleteDraft(draftId: number) {
  const db = await getDb();
  if (!db) throw new Error("Database not available");

  const result = await db.delete(drafts).where(eq(drafts.id, draftId));

  return result;
}
