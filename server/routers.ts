import { COOKIE_NAME } from "@shared/const";
import { getSessionCookieOptions } from "./_core/cookies";
import { systemRouter } from "./_core/systemRouter";
import { publicProcedure, protectedProcedure, router } from "./_core/trpc";
import { z } from "zod";
import {
  createReport,
  getReportById,
  getReportsByUserId,
  updateReport,
  deleteReport,
  addPhoto,
  getPhotosByReportId,
  updatePhoto,
  deletePhoto,
  saveDraft,
  getDraftByReportId,
  getDraftsByUserId,
  deleteDraft,
} from "./db";
import { notifyOwner } from "./_core/notification";
import { TRPCError } from "@trpc/server";

// Validation schemas
const checklistSchema = z.object({
  centralReady: z.boolean().nullable().optional().default(null),
  masonryReady: z.boolean().nullable().optional().default(null),
  tankReady: z.boolean().nullable().optional().default(null),
  networkExists: z.boolean().nullable().optional().default(null),
  pointsQuantityChange: z.boolean().nullable().optional().default(null),
  truckAccess: z.boolean().nullable().optional().default(null),
  extintorValid: z.boolean().nullable().optional().default(null),
  cellSignal: z.string().nullable().optional().default(null),
  cellOperator: z.string().nullable().optional().default(null),
  cellWifi: z.boolean().nullable().optional().default(null),
});

const quantitativesSchema = z.object({
  tankQuantity: z.string().optional().default(""),
  networkBitola: z.string().optional().default(""),
  networkMeterage: z.string().optional().default(""),
  pointsQuantity: z.string().optional().default(""),
  shelterQuantity: z.string().optional().default(""),
  resistantWallArea: z.string().optional().default(""),
  gateArea: z.string().optional().default(""),
  slabArea: z.string().optional().default(""),
});

const reportInputSchema = z.object({
  projectistName: z.string().optional(),
  projectistCompany: z.string().optional(),
  projectistPhone: z.string().optional(),
  clientCompanyName: z.string().optional(),
  clientAddress: z.string().optional(),
  clientCity: z.string().optional(),
  clientState: z.string().optional(),
  clientZipCode: z.string().optional(),
  clientNeighborhood: z.string().optional(),
  clientSANumber: z.string().optional(),
  clientARTNumber: z.string().optional(),
  unitLocation: z.string().optional(),
  isNewClient: z.enum(["yes", "no"]).optional(),
  isAsBuilt: z.enum(["yes", "no"]).optional(),
  isAdequation: z.enum(["yes", "no"]).optional(),
  checklist: checklistSchema.optional(),
  quantitatives: quantitativesSchema.optional(),
  logisticDifficulty: z.number().min(0).max(10).optional(),
  assemblyDifficulty: z.number().min(0).max(10).optional(),
  observations: z.string().optional(),
  clientActions: z.string().optional(),
  supergasActions: z.string().optional(),
  status: z.enum(["draft", "completed"]).optional(),
});

export const appRouter = router({
  system: systemRouter,
  auth: router({
    me: publicProcedure.query(opts => opts.ctx.user),
    logout: publicProcedure.mutation(({ ctx }) => {
      const cookieOptions = getSessionCookieOptions(ctx.req);
      ctx.res.clearCookie(COOKIE_NAME, { ...cookieOptions, maxAge: -1 });
      return {
        success: true,
      } as const;
    }),
  }),

  // Reports Router
  reports: router({
    // Create a new report
    create: protectedProcedure
      .input(reportInputSchema)
      .mutation(async ({ ctx, input }) => {
        const result = await createReport(ctx.user.id, {
          ...input,
        });

        // Notify owner
        await notifyOwner({
          title: "Novo Relatório de Visita Técnica",
          content: `Um novo relatório foi criado por ${ctx.user.name || "um usuário"}. Cliente: ${input.clientCompanyName || "Não informado"}`,
        });

        return { success: true };
      }),

    // Get report by ID
    getById: protectedProcedure
      .input(z.object({ id: z.number() }))
      .query(async ({ ctx, input }) => {
        const report = await getReportById(input.id);

        if (!report) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Relatório não encontrado",
          });
        }

        // Check authorization
        if (report.userId !== ctx.user.id && ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Você não tem permissão para acessar este relatório",
          });
        }

        // Get photos
        const reportPhotos = await getPhotosByReportId(report.id);

        return {
          ...report,
          photos: reportPhotos,
        };
      }),

    // List reports for current user
    list: protectedProcedure
      .input(
        z.object({
          limit: z.number().default(50),
          offset: z.number().default(0),
        })
      )
      .query(async ({ ctx, input }) => {
        const userReports = await getReportsByUserId(ctx.user.id, input.limit, input.offset);
        return userReports;
      }),

    // Update report
    update: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          data: reportInputSchema,
        })
      )
      .mutation(async ({ ctx, input }) => {
        const report = await getReportById(input.id);

        if (!report) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Relatório não encontrado",
          });
        }

        // Check authorization
        if (report.userId !== ctx.user.id && ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Você não tem permissão para editar este relatório",
          });
        }

        await updateReport(input.id, input.data);

        // If status changed to completed, notify owner
        if (input.data.status === "completed" && report.status !== "completed") {
          await notifyOwner({
            title: "Relatório de Visita Técnica Finalizado",
            content: `O relatório de visita técnica para ${input.data.clientCompanyName || report.clientCompanyName || "cliente"} foi finalizado.`,
          });
        }

        return { success: true };
      }),

    // Delete report
    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        const report = await getReportById(input.id);

        if (!report) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Relatório não encontrado",
          });
        }

        // Check authorization
        if (report.userId !== ctx.user.id && ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Você não tem permissão para deletar este relatório",
          });
        }

        await deleteReport(input.id);

        return { success: true };
      }),
  }),

  // Photos Router
  photos: router({
    // Add photo to report
    add: protectedProcedure
      .input(
        z.object({
          reportId: z.number(),
          photoUrl: z.string(),
          photoKey: z.string(),
          caption: z.string().optional(),
          order: z.number().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        const report = await getReportById(input.reportId);

        if (!report) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Relatório não encontrado",
          });
        }

        // Check authorization
        if (report.userId !== ctx.user.id && ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Você não tem permissão para adicionar fotos a este relatório",
          });
        }

        await addPhoto(input.reportId, {
          photoUrl: input.photoUrl,
          photoKey: input.photoKey,
          caption: input.caption || null,
          order: input.order || 0,
        });

        return { success: true };
      }),

    // Get photos for report
    getByReportId: protectedProcedure
      .input(z.object({ reportId: z.number() }))
      .query(async ({ ctx, input }) => {
        const report = await getReportById(input.reportId);

        if (!report) {
          throw new TRPCError({
            code: "NOT_FOUND",
            message: "Relatório não encontrado",
          });
        }

        // Check authorization
        if (report.userId !== ctx.user.id && ctx.user.role !== "admin") {
          throw new TRPCError({
            code: "FORBIDDEN",
            message: "Você não tem permissão para acessar as fotos deste relatório",
          });
        }

        return await getPhotosByReportId(input.reportId);
      }),

    // Update photo
    update: protectedProcedure
      .input(
        z.object({
          id: z.number(),
          caption: z.string().optional(),
          order: z.number().optional(),
        })
      )
      .mutation(async ({ ctx, input }) => {
        await updatePhoto(input.id, {
          caption: input.caption,
          order: input.order,
        });

        return { success: true };
      }),

    // Delete photo
    delete: protectedProcedure
      .input(z.object({ id: z.number() }))
      .mutation(async ({ ctx, input }) => {
        await deletePhoto(input.id);
        return { success: true };
      }),
  }),

  // Drafts Router
  drafts: router({
    // Save draft
    save: protectedProcedure
      .input(
        z.object({
          reportId: z.number().optional(),
          draftData: z.record(z.string(), z.unknown()),
        })
      )
      .mutation(async ({ ctx, input }) => {
        await saveDraft(ctx.user.id, input.reportId ?? null, input.draftData);
        return { success: true };
      }),

    // Get draft by report ID
    getByReportId: protectedProcedure
      .input(z.object({ reportId: z.number() }))
      .query(async ({ input }) => {
        const draft = await getDraftByReportId(input.reportId);
        return draft || null;
      }),

    // Get all drafts for user
    list: protectedProcedure.query(async ({ ctx }) => {
      return await getDraftsByUserId(ctx.user.id);
    }),

    // Delete draft
    delete: protectedProcedure
      .input(z.object({ draftId: z.number() }))
      .mutation(async ({ ctx, input }) => {
        await deleteDraft(input.draftId);

        return { success: true };
      }),
  }),
});

export type AppRouter = typeof appRouter;
