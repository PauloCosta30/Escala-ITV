import { useState, useEffect } from "react";
import { useRoute, useLocation } from "wouter";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Slider } from "@/components/ui/slider";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { trpc } from "@/lib/trpc";
import { toast } from "sonner";
import { Loader2, Save, ArrowLeft } from "lucide-react";

interface FormData {
  projectistName?: string | null;
  projectistCompany?: string | null;
  projectistPhone?: string | null;
  clientCompanyName?: string | null;
  clientAddress?: string | null;
  clientCity?: string | null;
  clientState?: string | null;
  clientZipCode?: string | null;
  clientNeighborhood?: string | null;
  clientSANumber?: string | null;
  clientARTNumber?: string | null;
  unitLocation?: string | null;
  isNewClient?: "yes" | "no" | null;
  isAsBuilt?: "yes" | "no" | null;
  isAdequation?: "yes" | "no" | null;
  checklist?: {
    centralReady?: boolean | null;
    masonryReady?: boolean | null;
    tankReady?: boolean | null;
    networkExists?: boolean | null;
    pointsQuantityChange?: boolean | null;
    truckAccess?: boolean | null;
    extintorValid?: boolean | null;
    cellSignal?: string | null;
    cellOperator?: string | null;
    cellWifi?: boolean | null;
  } | null;
  quantitatives?: {
    tankQuantity?: string;
    networkBitola?: string;
    networkMeterage?: string;
    pointsQuantity?: string;
    shelterQuantity?: string;
    resistantWallArea?: string;
    gateArea?: string;
    slabArea?: string;
  } | null;
  logisticDifficulty?: number | null;
  assemblyDifficulty?: number | null;
  observations?: string | null;
  clientActions?: string | null;
  supergasActions?: string | null;
  status?: "draft" | "completed" | null;
  photos?: any[];
  [key: string]: any;
}

export default function ReportForm() {
  const [match, params] = useRoute("/reports/:id");
  const [, navigate] = useLocation();
  const reportId = match && params?.id ? parseInt(params.id) : null;

  const [formData, setFormData] = useState<FormData>({
    isNewClient: "yes",
    isAsBuilt: "no",
    isAdequation: "no",
    logisticDifficulty: 5,
    assemblyDifficulty: 5,
    checklist: {},
    quantitatives: {},
  });

  const [isSaving, setIsSaving] = useState(false);
  const [autoSaveTimer, setAutoSaveTimer] = useState<NodeJS.Timeout | null>(null);

  const createMutation = trpc.reports.create.useMutation();
  const updateMutation = trpc.reports.update.useMutation();
  const getReportQuery = trpc.reports.getById.useQuery(
    { id: reportId! },
    { enabled: !!reportId }
  );
  const saveDraftMutation = trpc.drafts.save.useMutation();

  // Load report data if editing
  useEffect(() => {
    if (getReportQuery.data) {
      setFormData(getReportQuery.data);
    }
  }, [getReportQuery.data]);

  // Auto-save draft
  useEffect(() => {
    if (autoSaveTimer) clearTimeout(autoSaveTimer);

    const timer = setTimeout(async () => {
      try {
        await saveDraftMutation.mutateAsync({
          reportId: reportId ?? undefined,
          draftData: formData as Record<string, unknown>,
        });
      } catch (error) {
        console.error("Auto-save failed:", error);
      }
    }, 3000);

    setAutoSaveTimer(timer);

    return () => clearTimeout(timer);
  }, [formData, reportId, saveDraftMutation]);

  const handleInputChange = (field: string, value: any) => {
    setFormData((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleChecklistChange = (field: string, value: boolean | null | string) => {
    setFormData((prev) => ({
      ...prev,
      checklist: {
        ...prev.checklist,
        [field]: value,
      },
    }));
  };

  const handleQuantitativesChange = (field: string, value: string) => {
    setFormData((prev) => ({
      ...prev,
      quantitatives: {
        ...prev.quantitatives,
        [field]: value,
      },
    }));
  };

  const handleSubmit = async (status: "draft" | "completed") => {
    setIsSaving(true);
    try {
      const payload: any = {
        ...formData,
        status,
      };

      if (reportId) {
        await updateMutation.mutateAsync({
          id: reportId,
          data: payload,
        });
        toast.success("Relatório atualizado com sucesso!");
      } else {
        await createMutation.mutateAsync(payload);
        toast.success("Relatório criado com sucesso!");
      }

      navigate("/reports");
    } catch (error: any) {
      toast.error(error.message || "Erro ao salvar relatório");
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 py-8">
      <div className="max-w-4xl mx-auto px-4">
        <div className="flex items-center gap-4 mb-8">
          <Button
            variant="ghost"
            size="icon"
            onClick={() => navigate("/reports")}
          >
            <ArrowLeft className="w-4 h-4" />
          </Button>
          <div>
            <h1 className="text-3xl font-bold">
              {reportId ? "Editar Relatório" : "Novo Relatório"}
            </h1>
            <p className="text-gray-600">Preencha os dados da visita técnica</p>
          </div>
        </div>

        <Tabs defaultValue="projetista" className="w-full">
          <TabsList className="grid w-full grid-cols-5">
            <TabsTrigger value="projetista">Projetista</TabsTrigger>
            <TabsTrigger value="cliente">Cliente</TabsTrigger>
            <TabsTrigger value="central">Central</TabsTrigger>
            <TabsTrigger value="checklist">Checklist</TabsTrigger>
            <TabsTrigger value="observacoes">Observações</TabsTrigger>
          </TabsList>

          {/* TAB: PROJETISTA */}
          <TabsContent value="projetista">
            <Card>
              <CardHeader>
                <CardTitle>Dados do Projetista</CardTitle>
                <CardDescription>Informações do profissional responsável</CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <Label htmlFor="projectistName">Nome</Label>
                    <Input
                      id="projectistName"
                      value={formData.projectistName || ""}
                      onChange={(e) => handleInputChange("projectistName", e.target.value)}
                      placeholder="Nome do projetista"
                    />
                  </div>
                  <div>
                    <Label htmlFor="projectistPhone">Telefone</Label>
                    <Input
                      id="projectistPhone"
                      value={formData.projectistPhone || ""}
                      onChange={(e) => handleInputChange("projectistPhone", e.target.value)}
                      placeholder="(11) 99999-9999"
                    />
                  </div>
                </div>
                <div>
                  <Label htmlFor="projectistCompany">Empresa</Label>
                  <Input
                    id="projectistCompany"
                    value={formData.projectistCompany || ""}
                    onChange={(e) => handleInputChange("projectistCompany", e.target.value)}
                    placeholder="Nome da empresa"
                  />
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* TAB: CLIENTE */}
          <TabsContent value="cliente">
            <Card>
              <CardHeader>
                <CardTitle>Dados do Cliente</CardTitle>
                <CardDescription>Informações da empresa cliente</CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div>
                  <Label htmlFor="clientCompanyName">Razão Social</Label>
                  <Input
                    id="clientCompanyName"
                    value={formData.clientCompanyName || ""}
                    onChange={(e) => handleInputChange("clientCompanyName", e.target.value)}
                    placeholder="Nome da empresa"
                  />
                </div>

                <div>
                  <Label htmlFor="clientAddress">Endereço</Label>
                  <Input
                    id="clientAddress"
                    value={formData.clientAddress || ""}
                    onChange={(e) => handleInputChange("clientAddress", e.target.value)}
                    placeholder="Rua, número"
                  />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div>
                    <Label htmlFor="clientCity">Cidade</Label>
                    <Input
                      id="clientCity"
                      value={formData.clientCity || ""}
                      onChange={(e) => handleInputChange("clientCity", e.target.value)}
                      placeholder="Cidade"
                    />
                  </div>
                  <div>
                    <Label htmlFor="clientState">Estado</Label>
                    <Input
                      id="clientState"
                      value={formData.clientState || ""}
                      onChange={(e) => handleInputChange("clientState", e.target.value)}
                      placeholder="SP"
                      maxLength={2}
                    />
                  </div>
                  <div>
                    <Label htmlFor="clientZipCode">CEP</Label>
                    <Input
                      id="clientZipCode"
                      value={formData.clientZipCode || ""}
                      onChange={(e) => handleInputChange("clientZipCode", e.target.value)}
                      placeholder="00000-000"
                    />
                  </div>
                </div>

                <div>
                  <Label htmlFor="clientNeighborhood">Bairro</Label>
                  <Input
                    id="clientNeighborhood"
                    value={formData.clientNeighborhood || ""}
                    onChange={(e) => handleInputChange("clientNeighborhood", e.target.value)}
                    placeholder="Bairro"
                  />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <Label htmlFor="clientSANumber">Número SA</Label>
                    <Input
                      id="clientSANumber"
                      value={formData.clientSANumber || ""}
                      onChange={(e) => handleInputChange("clientSANumber", e.target.value)}
                      placeholder="SA-000000"
                    />
                  </div>
                  <div>
                    <Label htmlFor="clientARTNumber">Número ART</Label>
                    <Input
                      id="clientARTNumber"
                      value={formData.clientARTNumber || ""}
                      onChange={(e) => handleInputChange("clientARTNumber", e.target.value)}
                      placeholder="ART-000000"
                    />
                  </div>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* TAB: CENTRAL */}
          <TabsContent value="central">
            <Card>
              <CardHeader>
                <CardTitle>Informações da Central</CardTitle>
                <CardDescription>Dados da central de gás</CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div>
                  <Label htmlFor="unitLocation">Localização da Unidade</Label>
                  <Input
                    id="unitLocation"
                    value={formData.unitLocation || ""}
                    onChange={(e) => handleInputChange("unitLocation", e.target.value)}
                    placeholder="Localização da unidade"
                  />
                </div>

                <div className="space-y-4">
                  <Label>Tipo de Projeto</Label>
                  <div className="space-y-3">
                    <div className="flex items-center space-x-2">
                      <Checkbox
                        id="isNewClient"
                        checked={formData.isNewClient === "yes"}
                        onCheckedChange={(checked) =>
                          handleInputChange("isNewClient", checked ? "yes" : "no")
                        }
                      />
                      <Label htmlFor="isNewClient" className="font-normal cursor-pointer">
                        Novo Cliente
                      </Label>
                    </div>
                    <div className="flex items-center space-x-2">
                      <Checkbox
                        id="isAsBuilt"
                        checked={formData.isAsBuilt === "yes"}
                        onCheckedChange={(checked) =>
                          handleInputChange("isAsBuilt", checked ? "yes" : "no")
                        }
                      />
                      <Label htmlFor="isAsBuilt" className="font-normal cursor-pointer">
                        As Built
                      </Label>
                    </div>
                    <div className="flex items-center space-x-2">
                      <Checkbox
                        id="isAdequation"
                        checked={formData.isAdequation === "yes"}
                        onCheckedChange={(checked) =>
                          handleInputChange("isAdequation", checked ? "yes" : "no")
                        }
                      />
                      <Label htmlFor="isAdequation" className="font-normal cursor-pointer">
                        Adequação
                      </Label>
                    </div>
                  </div>
                </div>

                <div className="border-t pt-6">
                  <h3 className="font-semibold mb-4">Quantitativos/Escopo</h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <Label htmlFor="tankQuantity">Quantidade e Capacidade dos Tanques</Label>
                      <Input
                        id="tankQuantity"
                        value={formData.quantitatives?.tankQuantity || ""}
                        onChange={(e) => handleQuantitativesChange("tankQuantity", e.target.value)}
                        placeholder="Ex: 2x 13kg"
                      />
                    </div>
                    <div>
                      <Label htmlFor="networkBitola">Bitola da Rede</Label>
                      <Input
                        id="networkBitola"
                        value={formData.quantitatives?.networkBitola || ""}
                        onChange={(e) => handleQuantitativesChange("networkBitola", e.target.value)}
                        placeholder="Ex: 20mm"
                      />
                    </div>
                    <div>
                      <Label htmlFor="networkMeterage">Metragem da Rede</Label>
                      <Input
                        id="networkMeterage"
                        value={formData.quantitatives?.networkMeterage || ""}
                        onChange={(e) => handleQuantitativesChange("networkMeterage", e.target.value)}
                        placeholder="Ex: 25m"
                      />
                    </div>
                    <div>
                      <Label htmlFor="pointsQuantity">Quantidade de Pontos</Label>
                      <Input
                        id="pointsQuantity"
                        value={formData.quantitatives?.pointsQuantity || ""}
                        onChange={(e) => handleQuantitativesChange("pointsQuantity", e.target.value)}
                        placeholder="Ex: 3"
                      />
                    </div>
                    <div>
                      <Label htmlFor="shelterQuantity">Quantidade e Tipo de Abrigo</Label>
                      <Input
                        id="shelterQuantity"
                        value={formData.quantitatives?.shelterQuantity || ""}
                        onChange={(e) => handleQuantitativesChange("shelterQuantity", e.target.value)}
                        placeholder="Ex: 1 Nicho"
                      />
                    </div>
                    <div>
                      <Label htmlFor="resistantWallArea">m² de Parede Resistente ao Fogo</Label>
                      <Input
                        id="resistantWallArea"
                        value={formData.quantitatives?.resistantWallArea || ""}
                        onChange={(e) => handleQuantitativesChange("resistantWallArea", e.target.value)}
                        placeholder="Ex: 5"
                      />
                    </div>
                    <div>
                      <Label htmlFor="gateArea">Portão m²</Label>
                      <Input
                        id="gateArea"
                        value={formData.quantitatives?.gateArea || ""}
                        onChange={(e) => handleQuantitativesChange("gateArea", e.target.value)}
                        placeholder="Ex: 2"
                      />
                    </div>
                    <div>
                      <Label htmlFor="slabArea">Laje m²</Label>
                      <Input
                        id="slabArea"
                        value={formData.quantitatives?.slabArea || ""}
                        onChange={(e) => handleQuantitativesChange("slabArea", e.target.value)}
                        placeholder="Ex: 10"
                      />
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* TAB: CHECKLIST */}
          <TabsContent value="checklist">
            <Card>
              <CardHeader>
                <CardTitle>Checklist de Verificação</CardTitle>
                <CardDescription>Questões de verificação da visita técnica</CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div className="space-y-4">
                  {[
                    { key: "centralReady", label: "Central está pronta de acordo com norma vigente?" },
                    { key: "masonryReady", label: "Alvenaria da central (portão, piso, reboco, telhado)?" },
                    { key: "tankReady", label: "Tanque pode ser enviado com gás?" },
                    { key: "networkExists", label: "Rede existente?" },
                    { key: "pointsQuantityChange", label: "Ponto de consumo precisa alterar quantidade?" },
                    { key: "truckAccess", label: "Caminhão Supergasbras abastece com facilidade?" },
                    { key: "extintorValid", label: "Central já possui extintor válido e pressurizado?" },
                  ].map(({ key, label }) => (
                    <div key={key} className="flex items-center space-x-4 p-3 bg-gray-50 rounded">
                      <div className="flex-1">
                        <Label className="font-normal">{label}</Label>
                      </div>
                      <div className="flex items-center space-x-2">
                        <Button
                          variant={
                            formData.checklist?.[key as keyof typeof formData.checklist] === true
                              ? "default"
                              : "outline"
                          }
                          size="sm"
                          onClick={() => handleChecklistChange(key, true)}
                        >
                          Sim
                        </Button>
                        <Button
                          variant={
                            formData.checklist?.[key as keyof typeof formData.checklist] === false
                              ? "default"
                              : "outline"
                          }
                          size="sm"
                          onClick={() => handleChecklistChange(key, false)}
                        >
                          Não
                        </Button>
                      </div>
                    </div>
                  ))}
                </div>

                <div className="border-t pt-6">
                  <h3 className="font-semibold mb-4">Sinal de Celular</h3>
                  <div className="space-y-4">
                    <div>
                      <Label htmlFor="cellOperator">Operadora</Label>
                      <Input
                        id="cellOperator"
                        value={(formData.checklist?.cellOperator as string) || ""}
                        onChange={(e) => handleChecklistChange("cellOperator", e.target.value)}
                        placeholder="Ex: Vivo, Claro, Tim"
                      />
                    </div>
                    <div className="flex items-center space-x-2">
                      <Checkbox
                        id="cellWifi"
                        checked={formData.checklist?.cellWifi || false}
                        onCheckedChange={(checked) =>
                          handleChecklistChange("cellWifi", checked as boolean)
                        }
                      />
                      <Label htmlFor="cellWifi" className="font-normal cursor-pointer">
                        WiFi disponível
                      </Label>
                    </div>
                  </div>
                </div>

                <div className="border-t pt-6">
                  <h3 className="font-semibold mb-4">Avaliação de Dificuldade</h3>
                  <div className="space-y-6">
                    <div>
                      <div className="flex justify-between mb-2">
                        <Label>Dificuldade para Logística: {formData.logisticDifficulty}/10</Label>
                      </div>
                      <Slider
                        value={[formData.logisticDifficulty || 5]}
                        onValueChange={(value) =>
                          handleInputChange("logisticDifficulty", value[0])
                        }
                        min={0}
                        max={10}
                        step={1}
                        className="w-full"
                      />
                    </div>
                    <div>
                      <div className="flex justify-between mb-2">
                        <Label>Dificuldade para Montagem: {formData.assemblyDifficulty}/10</Label>
                      </div>
                      <Slider
                        value={[formData.assemblyDifficulty || 5]}
                        onValueChange={(value) =>
                          handleInputChange("assemblyDifficulty", value[0])
                        }
                        min={0}
                        max={10}
                        step={1}
                        className="w-full"
                      />
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          {/* TAB: OBSERVAÇÕES */}
          <TabsContent value="observacoes">
            <Card>
              <CardHeader>
                <CardTitle>Observações e Ações</CardTitle>
                <CardDescription>Informações adicionais e plano de ação</CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div>
                  <Label htmlFor="observations">Observações Gerais</Label>
                  <Textarea
                    id="observations"
                    value={formData.observations || ""}
                    onChange={(e) => handleInputChange("observations", e.target.value)}
                    placeholder="Observações adicionais sobre a visita"
                    rows={4}
                  />
                </div>

                <div>
                  <Label htmlFor="clientActions">Ações do Cliente</Label>
                  <Textarea
                    id="clientActions"
                    value={formData.clientActions || ""}
                    onChange={(e) => handleInputChange("clientActions", e.target.value)}
                    placeholder="Ações que o cliente precisa executar"
                    rows={4}
                  />
                </div>

                <div>
                  <Label htmlFor="supergasActions">Ações da Supergasbras</Label>
                  <Textarea
                    id="supergasActions"
                    value={formData.supergasActions || ""}
                    onChange={(e) => handleInputChange("supergasActions", e.target.value)}
                    placeholder="Ações que a Supergasbras precisa executar"
                    rows={4}
                  />
                </div>
              </CardContent>
            </Card>
          </TabsContent>
        </Tabs>

        {/* Buttons */}
        <div className="flex gap-4 mt-8">
          <Button
            variant="outline"
            onClick={() => navigate("/reports")}
            disabled={isSaving}
          >
            Cancelar
          </Button>
          <Button
            variant="secondary"
            onClick={() => handleSubmit("draft")}
            disabled={isSaving}
          >
            {isSaving && <Loader2 className="w-4 h-4 mr-2 animate-spin" />}
            Salvar como Rascunho
          </Button>
          <Button
            onClick={() => handleSubmit("completed")}
            disabled={isSaving}
          >
            {isSaving && <Loader2 className="w-4 h-4 mr-2 animate-spin" />}
            Finalizar Relatório
          </Button>
        </div>
      </div>
    </div>
  );
}
