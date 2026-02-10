import { useState } from "react";
import { useLocation } from "wouter";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";
import { trpc } from "@/lib/trpc";
import { toast } from "sonner";
import { Plus, FileText, Edit2, Trash2, Eye, Download, Loader2 } from "lucide-react";

export default function ReportsList() {
  const [, navigate] = useLocation();
  const [deleteId, setDeleteId] = useState<number | null>(null);
  const [exportingId, setExportingId] = useState<number | null>(null);

  const listQuery = trpc.reports.list.useQuery({ limit: 50, offset: 0 });
  const deleteMutation = trpc.reports.delete.useMutation();
  const getReportQuery = trpc.reports.getById.useQuery(
    { id: exportingId! },
    { enabled: !!exportingId }
  );

  const handleDelete = async () => {
    if (!deleteId) return;

    try {
      await deleteMutation.mutateAsync({ id: deleteId });
      toast.success("Relatório deletado com sucesso!");
      listQuery.refetch();
      setDeleteId(null);
    } catch (error: any) {
      toast.error(error.message || "Erro ao deletar relatório");
    }
  };

  const handleExport = async (id: number) => {
    setExportingId(id);
    // PDF export will be handled in the next phase
    toast.info("Funcionalidade de exportação em desenvolvimento");
  };

  const formatDate = (date: Date) => {
    return new Date(date).toLocaleDateString("pt-BR", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });
  };

  const getStatusBadge = (status: string) => {
    return status === "completed" ? (
      <Badge className="bg-green-100 text-green-800">Finalizado</Badge>
    ) : (
      <Badge className="bg-yellow-100 text-yellow-800">Rascunho</Badge>
    );
  };

  return (
    <div className="min-h-screen bg-gray-50 py-8">
      <div className="max-w-6xl mx-auto px-4">
        <div className="flex items-center justify-between mb-8">
          <div>
            <h1 className="text-3xl font-bold">Relatórios de Visita Técnica</h1>
            <p className="text-gray-600">Gerencie seus relatórios de visita técnica</p>
          </div>
          <Button onClick={() => navigate("/reports/new")} size="lg">
            <Plus className="w-4 h-4 mr-2" />
            Novo Relatório
          </Button>
        </div>

        {listQuery.isLoading ? (
          <Card>
            <CardContent className="flex items-center justify-center py-12">
              <Loader2 className="w-8 h-8 animate-spin text-gray-400" />
            </CardContent>
          </Card>
        ) : listQuery.data && listQuery.data.length > 0 ? (
          <Card>
            <CardHeader>
              <CardTitle>Seus Relatórios</CardTitle>
              <CardDescription>
                Total de {listQuery.data.length} relatório(s)
              </CardDescription>
            </CardHeader>
            <CardContent>
              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Cliente</TableHead>
                      <TableHead>Localização</TableHead>
                      <TableHead>Data</TableHead>
                      <TableHead>Status</TableHead>
                      <TableHead className="text-right">Ações</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {listQuery.data.map((report: any) => (
                      <TableRow key={report.id}>
                        <TableCell className="font-medium">
                          {report.clientCompanyName || "Sem nome"}
                        </TableCell>
                        <TableCell>{report.unitLocation || "-"}</TableCell>
                        <TableCell>{formatDate(report.createdAt)}</TableCell>
                        <TableCell>{getStatusBadge(report.status)}</TableCell>
                        <TableCell className="text-right">
                          <div className="flex items-center justify-end gap-2">
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => navigate(`/reports/${report.id}`)}
                              title="Visualizar"
                            >
                              <Eye className="w-4 h-4" />
                            </Button>
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => navigate(`/reports/${report.id}/edit`)}
                              title="Editar"
                            >
                              <Edit2 className="w-4 h-4" />
                            </Button>
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => handleExport(report.id)}
                              title="Exportar PDF"
                            >
                              <Download className="w-4 h-4" />
                            </Button>
                            <Button
                              variant="ghost"
                              size="sm"
                              onClick={() => setDeleteId(report.id)}
                              title="Deletar"
                            >
                              <Trash2 className="w-4 h-4 text-red-500" />
                            </Button>
                          </div>
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
            </CardContent>
          </Card>
        ) : (
          <Card>
            <CardContent className="flex flex-col items-center justify-center py-12">
              <FileText className="w-16 h-16 text-gray-300 mb-4" />
              <h3 className="text-lg font-semibold text-gray-600 mb-2">
                Nenhum relatório encontrado
              </h3>
              <p className="text-gray-500 mb-6">
                Comece criando seu primeiro relatório de visita técnica
              </p>
              <Button onClick={() => navigate("/reports/new")}>
                <Plus className="w-4 h-4 mr-2" />
                Criar Novo Relatório
              </Button>
            </CardContent>
          </Card>
        )}

        {/* Delete Confirmation Dialog */}
        <AlertDialog open={deleteId !== null} onOpenChange={() => setDeleteId(null)}>
          <AlertDialogContent>
            <AlertDialogTitle>Deletar Relatório</AlertDialogTitle>
            <AlertDialogDescription>
              Tem certeza que deseja deletar este relatório? Esta ação não pode ser desfeita.
            </AlertDialogDescription>
            <div className="flex gap-4 justify-end">
              <AlertDialogCancel>Cancelar</AlertDialogCancel>
              <AlertDialogAction
                onClick={handleDelete}
                className="bg-red-600 hover:bg-red-700"
              >
                Deletar
              </AlertDialogAction>
            </div>
          </AlertDialogContent>
        </AlertDialog>
      </div>
    </div>
  );
}
