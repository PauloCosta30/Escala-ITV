import { useEffect, useRef } from "react";
import { useRoute, useLocation } from "wouter";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { trpc } from "@/lib/trpc";
import { toast } from "sonner";
import { ArrowLeft, Download, Edit2, Loader2, Printer } from "lucide-react";
import html2pdf from "html2pdf.js";

export default function ReportView() {
  const [match, params] = useRoute("/reports/:id");
  const [, navigate] = useLocation();
  const reportId = match && params?.id ? parseInt(params.id) : null;
  const reportRef = useRef<HTMLDivElement>(null);

  const getReportQuery = trpc.reports.getById.useQuery(
    { id: reportId! },
    { enabled: !!reportId }
  );

  const handleExportPDF = async () => {
    if (!reportRef.current) return;

    try {
      const element = reportRef.current;
      const opt: any = {
        margin: 10,
        filename: `relatorio-${reportId}.pdf`,
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: "portrait", unit: "mm", format: "a4" },
      };

      (html2pdf() as any).set(opt).from(element).save();
      toast.success("Relatório exportado com sucesso!");
    } catch (error: any) {
      toast.error("Erro ao exportar relatório");
    }
  };

  const handlePrint = () => {
    window.print();
  };

  if (!reportId) return null;

  if (getReportQuery.isLoading) {
    return (
      <div className="min-h-screen bg-gray-50 py-8 flex items-center justify-center">
        <Loader2 className="w-8 h-8 animate-spin text-gray-400" />
      </div>
    );
  }

  if (!getReportQuery.data) {
    return (
      <div className="min-h-screen bg-gray-50 py-8">
        <div className="max-w-4xl mx-auto px-4">
          <Button
            variant="ghost"
            onClick={() => navigate("/reports")}
            className="mb-4"
          >
            <ArrowLeft className="w-4 h-4 mr-2" />
            Voltar
          </Button>
          <Card>
            <CardContent className="py-12 text-center">
              <p className="text-gray-500">Relatório não encontrado</p>
            </CardContent>
          </Card>
        </div>
      </div>
    );
  }

  const report = getReportQuery.data;

  return (
    <div className="min-h-screen bg-gray-50 py-8">
      <div className="max-w-4xl mx-auto px-4">
        <div className="flex items-center justify-between mb-8">
          <Button
            variant="ghost"
            onClick={() => navigate("/reports")}
          >
            <ArrowLeft className="w-4 h-4 mr-2" />
            Voltar
          </Button>
          <div className="flex gap-2">
            <Button
              variant="outline"
              onClick={handlePrint}
            >
              <Printer className="w-4 h-4 mr-2" />
              Imprimir
            </Button>
            <Button
              variant="outline"
              onClick={handleExportPDF}
            >
              <Download className="w-4 h-4 mr-2" />
              Exportar PDF
            </Button>
            <Button
              onClick={() => navigate(`/reports/${reportId}/edit`)}
            >
              <Edit2 className="w-4 h-4 mr-2" />
              Editar
            </Button>
          </div>
        </div>

        <div ref={reportRef} className="bg-white p-8 rounded-lg shadow">
          {/* Header */}
          <div className="mb-8 border-b pb-6">
            <div className="flex items-center justify-between mb-4">
              <h1 className="text-3xl font-bold">RELATÓRIO DE VISITA TÉCNICA</h1>
              <Badge variant={report.status === "completed" ? "default" : "secondary"}>
                {report.status === "completed" ? "Finalizado" : "Rascunho"}
              </Badge>
            </div>
            <p className="text-gray-600">
              Data: {new Date(report.createdAt).toLocaleDateString("pt-BR")}
            </p>
          </div>

          {/* Projetista */}
          <div className="mb-8">
            <h2 className="text-xl font-bold mb-4">DADOS DO PROJETISTA</h2>
            <div className="grid grid-cols-2 gap-4 text-sm">
              <div>
                <p className="font-semibold">Nome</p>
                <p className="text-gray-600">{report.projectistName || "-"}</p>
              </div>
              <div>
                <p className="font-semibold">Telefone</p>
                <p className="text-gray-600">{report.projectistPhone || "-"}</p>
              </div>
              <div className="col-span-2">
                <p className="font-semibold">Empresa</p>
                <p className="text-gray-600">{report.projectistCompany || "-"}</p>
              </div>
            </div>
          </div>

          {/* Cliente */}
          <div className="mb-8">
            <h2 className="text-xl font-bold mb-4">DADOS DO CLIENTE</h2>
            <div className="grid grid-cols-2 gap-4 text-sm">
              <div className="col-span-2">
                <p className="font-semibold">Razão Social</p>
                <p className="text-gray-600">{report.clientCompanyName || "-"}</p>
              </div>
              <div className="col-span-2">
                <p className="font-semibold">Endereço</p>
                <p className="text-gray-600">{report.clientAddress || "-"}</p>
              </div>
              <div>
                <p className="font-semibold">Cidade</p>
                <p className="text-gray-600">{report.clientCity || "-"}</p>
              </div>
              <div>
                <p className="font-semibold">Estado</p>
                <p className="text-gray-600">{report.clientState || "-"}</p>
              </div>
              <div>
                <p className="font-semibold">CEP</p>
                <p className="text-gray-600">{report.clientZipCode || "-"}</p>
              </div>
              <div>
                <p className="font-semibold">Bairro</p>
                <p className="text-gray-600">{report.clientNeighborhood || "-"}</p>
              </div>
              <div>
                <p className="font-semibold">Número SA</p>
                <p className="text-gray-600">{report.clientSANumber || "-"}</p>
              </div>
              <div>
                <p className="font-semibold">Número ART</p>
                <p className="text-gray-600">{report.clientARTNumber || "-"}</p>
              </div>
            </div>
          </div>

          {/* Central */}
          <div className="mb-8">
            <h2 className="text-xl font-bold mb-4">INFORMAÇÕES DA CENTRAL</h2>
            <div className="grid grid-cols-2 gap-4 text-sm mb-4">
              <div className="col-span-2">
                <p className="font-semibold">Localização da Unidade</p>
                <p className="text-gray-600">{report.unitLocation || "-"}</p>
              </div>
              <div>
                <p className="font-semibold">Novo Cliente</p>
                <p className="text-gray-600">{report.isNewClient === "yes" ? "Sim" : "Não"}</p>
              </div>
              <div>
                <p className="font-semibold">As Built</p>
                <p className="text-gray-600">{report.isAsBuilt === "yes" ? "Sim" : "Não"}</p>
              </div>
              <div>
                <p className="font-semibold">Adequação</p>
                <p className="text-gray-600">{report.isAdequation === "yes" ? "Sim" : "Não"}</p>
              </div>
            </div>

            {report.quantitatives && (
              <div className="mt-4">
                <p className="font-semibold mb-2">Quantitativos/Escopo</p>
                <div className="grid grid-cols-2 gap-2 text-sm">
                  {Object.entries(report.quantitatives).map(([key, value]) => (
                    <div key={key}>
                      <p className="font-semibold text-xs">{key}</p>
                      <p className="text-gray-600">{value || "-"}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Checklist */}
          {report.checklist && (
            <div className="mb-8">
              <h2 className="text-xl font-bold mb-4">CHECKLIST DE VERIFICAÇÃO</h2>
              <div className="text-sm space-y-2">
                {Object.entries(report.checklist).map(([key, value]) => (
                  <div key={key} className="flex justify-between">
                    <span className="font-semibold">{key}:</span>
                    <span className="text-gray-600">
                      {typeof value === "boolean" ? (value ? "Sim" : "Não") : value || "-"}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Dificuldade */}
          <div className="mb-8">
            <h2 className="text-xl font-bold mb-4">AVALIAÇÃO DE DIFICULDADE</h2>
            <div className="grid grid-cols-2 gap-4 text-sm">
              <div>
                <p className="font-semibold">Logística: {report.logisticDifficulty}/10</p>
              </div>
              <div>
                <p className="font-semibold">Montagem: {report.assemblyDifficulty}/10</p>
              </div>
            </div>
          </div>

          {/* Observações */}
          {report.observations && (
            <div className="mb-8">
              <h2 className="text-xl font-bold mb-4">OBSERVAÇÕES</h2>
              <p className="text-sm text-gray-600 whitespace-pre-wrap">{report.observations}</p>
            </div>
          )}

          {/* Ações */}
          {(report.clientActions || report.supergasActions) && (
            <div className="mb-8">
              <h2 className="text-xl font-bold mb-4">AÇÕES</h2>
              {report.clientActions && (
                <div className="mb-4">
                  <p className="font-semibold text-sm mb-2">Ações do Cliente</p>
                  <p className="text-sm text-gray-600 whitespace-pre-wrap">
                    {report.clientActions}
                  </p>
                </div>
              )}
              {report.supergasActions && (
                <div>
                  <p className="font-semibold text-sm mb-2">Ações da Supergasbras</p>
                  <p className="text-sm text-gray-600 whitespace-pre-wrap">
                    {report.supergasActions}
                  </p>
                </div>
              )}
            </div>
          )}

          {/* Fotos */}
          {report.photos && report.photos.length > 0 && (
            <div className="mb-8">
              <h2 className="text-xl font-bold mb-4">FOTOS DA VISITA</h2>
              <div className="grid grid-cols-2 gap-4">
                {report.photos.map((photo: any) => (
                  <div key={photo.id}>
                    <img
                      src={photo.photoUrl}
                      alt={photo.caption || "Foto"}
                      className="w-full h-auto rounded border"
                    />
                    {photo.caption && (
                      <p className="text-xs text-gray-600 mt-2">{photo.caption}</p>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
