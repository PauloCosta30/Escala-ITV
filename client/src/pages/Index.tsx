import { useState, useRef } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Checkbox } from "@/components/ui/checkbox";
import { Slider } from "@/components/ui/slider";
import { toast } from "sonner";
import { Upload, X, Download, Loader2, Image as ImageIcon } from "lucide-react";
import html2pdf from "html2pdf.js";

interface Photo {
  id: string;
  file: File;
  dataUrl: string;
  observation: string;
}

interface FormData {
  // Projetista
  projectistName: string;
  projectistPhone: string;
  projectistCompany: string;

  // Cliente
  clientCompanyName: string;
  clientAddress: string;
  clientCity: string;
  clientState: string;
  clientZipCode: string;
  clientNeighborhood: string;
  clientSANumber: string;
  clientARTNumber: string;

  // Central
  unitLocation: string;
  isNewClient: boolean;
  isAsBuilt: boolean;
  isAdequation: boolean;

  // Quantitativos
  tankQuantity: string;
  networkBitola: string;
  networkMeterage: string;
  pointsQuantity: string;
  shelterQuantity: string;
  resistantWallArea: string;
  gateArea: string;
  slabArea: string;

  // Checklist
  centralReady: boolean;
  masonryReady: boolean;
  tankReady: boolean;
  networkExists: boolean;
  pointsQuantityChange: boolean;
  truckAccess: boolean;
  extintorValid: boolean;
  cellSignal: string;
  cellOperator: string;
  cellWifi: boolean;

  // Dificuldade
  logisticDifficulty: number;
  assemblyDifficulty: number;

  // Observações
  observations: string;
  clientActions: string;
  supergasActions: string;
}

export default function Index() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const reportRef = useRef<HTMLDivElement>(null);
  const [isExporting, setIsExporting] = useState(false);

  const [photos, setPhotos] = useState<Photo[]>([]);
  const [currentPhotoObservation, setCurrentPhotoObservation] = useState("");

  const [formData, setFormData] = useState<FormData>({
    projectistName: "",
    projectistPhone: "",
    projectistCompany: "",
    clientCompanyName: "",
    clientAddress: "",
    clientCity: "",
    clientState: "",
    clientZipCode: "",
    clientNeighborhood: "",
    clientSANumber: "",
    clientARTNumber: "",
    unitLocation: "",
    isNewClient: false,
    isAsBuilt: false,
    isAdequation: false,
    tankQuantity: "",
    networkBitola: "",
    networkMeterage: "",
    pointsQuantity: "",
    shelterQuantity: "",
    resistantWallArea: "",
    gateArea: "",
    slabArea: "",
    centralReady: false,
    masonryReady: false,
    tankReady: false,
    networkExists: false,
    pointsQuantityChange: false,
    truckAccess: false,
    extintorValid: false,
    cellSignal: "",
    cellOperator: "",
    cellWifi: false,
    logisticDifficulty: 5,
    assemblyDifficulty: 5,
    observations: "",
    clientActions: "",
    supergasActions: "",
  });

  const handleInputChange = (field: keyof FormData, value: any) => {
    setFormData((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!file.type.startsWith("image/")) {
      toast.error("Por favor, selecione uma imagem válida");
      return;
    }

    if (file.size > 5 * 1024 * 1024) {
      toast.error("A imagem deve ter menos de 5MB");
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      const dataUrl = event.target?.result as string;
      const newPhoto: Photo = {
        id: Date.now().toString(),
        file,
        dataUrl,
        observation: currentPhotoObservation,
      };
      setPhotos((prev) => [...prev, newPhoto]);
      setCurrentPhotoObservation("");
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
      toast.success("Foto adicionada com sucesso!");
    };
    reader.readAsDataURL(file);
  };

  const handleDeletePhoto = (photoId: string) => {
    setPhotos((prev) => prev.filter((p) => p.id !== photoId));
    toast.success("Foto removida");
  };

  const handleExportPDF = async () => {
    if (!reportRef.current) return;

    setIsExporting(true);
    try {
      const element = reportRef.current;
      const opt: any = {
        margin: 10,
        filename: `relatorio-visita-tecnica-${new Date().toISOString().split("T")[0]}.pdf`,
        image: { type: "jpeg", quality: 0.95 },
        html2canvas: { scale: 2, useCORS: true, logging: false },
        jsPDF: { orientation: "portrait", unit: "mm", format: "a4" },
      };

      const pdf = (html2pdf() as any);
      pdf.set(opt).from(element).save();
      toast.success("Relatório exportado com sucesso!");
    } catch (error: any) {
      toast.error("Erro ao exportar relatório");
      console.error(error);
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 py-8">
      <div className="max-w-5xl mx-auto px-4">
        {/* Header */}
        <div className="mb-8">
          <h1 className="text-4xl font-bold text-gray-900 mb-2">
            Relatório de Visita Técnica
          </h1>
          <p className="text-gray-600">Supergasbras - Preencha o formulário e exporte em PDF</p>
        </div>

        {/* Botão de Exportação */}
        <div className="mb-6 flex justify-end">
          <Button
            onClick={handleExportPDF}
            disabled={isExporting}
            size="lg"
            className="bg-green-600 hover:bg-green-700"
          >
            {isExporting ? (
              <>
                <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                Exportando...
              </>
            ) : (
              <>
                <Download className="w-4 h-4 mr-2" />
                Exportar PDF
              </>
            )}
          </Button>
        </div>

        {/* Formulário */}
        <div ref={reportRef} className="bg-white rounded-lg shadow-lg p-8">
          <Tabs defaultValue="projetista" className="w-full">
            <TabsList className="grid w-full grid-cols-5">
              <TabsTrigger value="projetista">Projetista</TabsTrigger>
              <TabsTrigger value="cliente">Cliente</TabsTrigger>
              <TabsTrigger value="central">Central</TabsTrigger>
              <TabsTrigger value="checklist">Checklist</TabsTrigger>
              <TabsTrigger value="observacoes">Observações</TabsTrigger>
            </TabsList>

            {/* Aba Projetista */}
            <TabsContent value="projetista" className="space-y-6 mt-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <Label htmlFor="projectistName">Nome do Projetista</Label>
                  <Input
                    id="projectistName"
                    value={formData.projectistName}
                    onChange={(e) => handleInputChange("projectistName", e.target.value)}
                    placeholder="Ex: Bruno Costa"
                  />
                </div>
                <div>
                  <Label htmlFor="projectistPhone">Telefone</Label>
                  <Input
                    id="projectistPhone"
                    value={formData.projectistPhone}
                    onChange={(e) => handleInputChange("projectistPhone", e.target.value)}
                    placeholder="Ex: (11) 92144-4173"
                  />
                </div>
              </div>
              <div>
                <Label htmlFor="projectistCompany">Empresa</Label>
                <Input
                  id="projectistCompany"
                  value={formData.projectistCompany}
                  onChange={(e) => handleInputChange("projectistCompany", e.target.value)}
                  placeholder="Ex: Empresa de Projetos"
                />
              </div>
            </TabsContent>

            {/* Aba Cliente */}
            <TabsContent value="cliente" className="space-y-6 mt-6">
              <div>
                <Label htmlFor="clientCompanyName">Razão Social</Label>
                <Input
                  id="clientCompanyName"
                  value={formData.clientCompanyName}
                  onChange={(e) => handleInputChange("clientCompanyName", e.target.value)}
                  placeholder="Ex: SKINAO PIZZARIA BREDA LTDA"
                />
              </div>
              <div>
                <Label htmlFor="clientAddress">Endereço</Label>
                <Input
                  id="clientAddress"
                  value={formData.clientAddress}
                  onChange={(e) => handleInputChange("clientAddress", e.target.value)}
                  placeholder="Ex: AVENIDA DA SAUDADE, 863"
                />
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <Label htmlFor="clientCity">Cidade</Label>
                  <Input
                    id="clientCity"
                    value={formData.clientCity}
                    onChange={(e) => handleInputChange("clientCity", e.target.value)}
                    placeholder="Ex: Miracatu"
                  />
                </div>
                <div>
                  <Label htmlFor="clientState">Estado</Label>
                  <Input
                    id="clientState"
                    value={formData.clientState}
                    onChange={(e) => handleInputChange("clientState", e.target.value)}
                    placeholder="Ex: SP"
                    maxLength={2}
                  />
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <Label htmlFor="clientZipCode">CEP</Label>
                  <Input
                    id="clientZipCode"
                    value={formData.clientZipCode}
                    onChange={(e) => handleInputChange("clientZipCode", e.target.value)}
                    placeholder="Ex: 11850-000"
                  />
                </div>
                <div>
                  <Label htmlFor="clientNeighborhood">Bairro</Label>
                  <Input
                    id="clientNeighborhood"
                    value={formData.clientNeighborhood}
                    onChange={(e) => handleInputChange("clientNeighborhood", e.target.value)}
                    placeholder="Ex: Vila Ubirajara"
                  />
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <Label htmlFor="clientSANumber">Número SA</Label>
                  <Input
                    id="clientSANumber"
                    value={formData.clientSANumber}
                    onChange={(e) => handleInputChange("clientSANumber", e.target.value)}
                    placeholder="Ex: 123456"
                  />
                </div>
                <div>
                  <Label htmlFor="clientARTNumber">Número ART</Label>
                  <Input
                    id="clientARTNumber"
                    value={formData.clientARTNumber}
                    onChange={(e) => handleInputChange("clientARTNumber", e.target.value)}
                    placeholder="Ex: 789456"
                  />
                </div>
              </div>
            </TabsContent>

            {/* Aba Central */}
            <TabsContent value="central" className="space-y-6 mt-6">
              <div>
                <Label htmlFor="unitLocation">Localização da Unidade</Label>
                <Input
                  id="unitLocation"
                  value={formData.unitLocation}
                  onChange={(e) => handleInputChange("unitLocation", e.target.value)}
                  placeholder="Ex: Mauá - SP"
                />
              </div>

              <div className="space-y-4">
                <h3 className="font-semibold">Informações da Central</h3>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="isNewClient"
                    checked={formData.isNewClient}
                    onCheckedChange={(checked) =>
                      handleInputChange("isNewClient", checked)
                    }
                  />
                  <Label htmlFor="isNewClient" className="font-normal cursor-pointer">
                    Novo Cliente
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="isAsBuilt"
                    checked={formData.isAsBuilt}
                    onCheckedChange={(checked) => handleInputChange("isAsBuilt", checked)}
                  />
                  <Label htmlFor="isAsBuilt" className="font-normal cursor-pointer">
                    As Built
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="isAdequation"
                    checked={formData.isAdequation}
                    onCheckedChange={(checked) => handleInputChange("isAdequation", checked)}
                  />
                  <Label htmlFor="isAdequation" className="font-normal cursor-pointer">
                    Adequação
                  </Label>
                </div>
              </div>

              <div className="mt-8">
                <h3 className="font-semibold mb-4">Quantitativos/Escopo</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <Label htmlFor="tankQuantity">Quantidade de Cilindros</Label>
                    <Input
                      id="tankQuantity"
                      value={formData.tankQuantity}
                      onChange={(e) => handleInputChange("tankQuantity", e.target.value)}
                      placeholder="Ex: 2x 13kg"
                    />
                  </div>
                  <div>
                    <Label htmlFor="networkBitola">Bitola da Rede</Label>
                    <Input
                      id="networkBitola"
                      value={formData.networkBitola}
                      onChange={(e) => handleInputChange("networkBitola", e.target.value)}
                      placeholder="Ex: 20mm"
                    />
                  </div>
                  <div>
                    <Label htmlFor="networkMeterage">Metragem da Rede</Label>
                    <Input
                      id="networkMeterage"
                      value={formData.networkMeterage}
                      onChange={(e) => handleInputChange("networkMeterage", e.target.value)}
                      placeholder="Ex: 25m"
                    />
                  </div>
                  <div>
                    <Label htmlFor="pointsQuantity">Quantidade de Pontos</Label>
                    <Input
                      id="pointsQuantity"
                      value={formData.pointsQuantity}
                      onChange={(e) => handleInputChange("pointsQuantity", e.target.value)}
                      placeholder="Ex: 3"
                    />
                  </div>
                  <div>
                    <Label htmlFor="shelterQuantity">Quantidade de Abrigos</Label>
                    <Input
                      id="shelterQuantity"
                      value={formData.shelterQuantity}
                      onChange={(e) => handleInputChange("shelterQuantity", e.target.value)}
                      placeholder="Ex: 1 Nicho"
                    />
                  </div>
                  <div>
                    <Label htmlFor="resistantWallArea">Área de Parede Resistente</Label>
                    <Input
                      id="resistantWallArea"
                      value={formData.resistantWallArea}
                      onChange={(e) => handleInputChange("resistantWallArea", e.target.value)}
                      placeholder="Ex: 5"
                    />
                  </div>
                  <div>
                    <Label htmlFor="gateArea">Área de Portão</Label>
                    <Input
                      id="gateArea"
                      value={formData.gateArea}
                      onChange={(e) => handleInputChange("gateArea", e.target.value)}
                      placeholder="Ex: 2"
                    />
                  </div>
                  <div>
                    <Label htmlFor="slabArea">Área de Laje</Label>
                    <Input
                      id="slabArea"
                      value={formData.slabArea}
                      onChange={(e) => handleInputChange("slabArea", e.target.value)}
                      placeholder="Ex: 10"
                    />
                  </div>
                </div>
              </div>
            </TabsContent>

            {/* Aba Checklist */}
            <TabsContent value="checklist" className="space-y-6 mt-6">
              <h3 className="font-semibold text-lg">Checklist de Verificação</h3>
              <div className="space-y-4">
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="centralReady"
                    checked={formData.centralReady}
                    onCheckedChange={(checked) => handleInputChange("centralReady", checked)}
                  />
                  <Label htmlFor="centralReady" className="font-normal cursor-pointer">
                    Central Pronta
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="masonryReady"
                    checked={formData.masonryReady}
                    onCheckedChange={(checked) => handleInputChange("masonryReady", checked)}
                  />
                  <Label htmlFor="masonryReady" className="font-normal cursor-pointer">
                    Alvenaria Pronta
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="tankReady"
                    checked={formData.tankReady}
                    onCheckedChange={(checked) => handleInputChange("tankReady", checked)}
                  />
                  <Label htmlFor="tankReady" className="font-normal cursor-pointer">
                    Cilindros Prontos
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="networkExists"
                    checked={formData.networkExists}
                    onCheckedChange={(checked) => handleInputChange("networkExists", checked)}
                  />
                  <Label htmlFor="networkExists" className="font-normal cursor-pointer">
                    Rede Existe
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="pointsQuantityChange"
                    checked={formData.pointsQuantityChange}
                    onCheckedChange={(checked) =>
                      handleInputChange("pointsQuantityChange", checked)
                    }
                  />
                  <Label
                    htmlFor="pointsQuantityChange"
                    className="font-normal cursor-pointer"
                  >
                    Quantidade de Pontos Alterada
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="truckAccess"
                    checked={formData.truckAccess}
                    onCheckedChange={(checked) => handleInputChange("truckAccess", checked)}
                  />
                  <Label htmlFor="truckAccess" className="font-normal cursor-pointer">
                    Acesso para Caminhão
                  </Label>
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="extintorValid"
                    checked={formData.extintorValid}
                    onCheckedChange={(checked) => handleInputChange("extintorValid", checked)}
                  />
                  <Label htmlFor="extintorValid" className="font-normal cursor-pointer">
                    Extintor Válido
                  </Label>
                </div>
              </div>

              <div className="mt-8 space-y-4">
                <div>
                  <Label htmlFor="cellSignal">Sinal de Celular</Label>
                  <Input
                    id="cellSignal"
                    value={formData.cellSignal}
                    onChange={(e) => handleInputChange("cellSignal", e.target.value)}
                    placeholder="Ex: Bom, Fraco, Nenhum"
                  />
                </div>
                <div>
                  <Label htmlFor="cellOperator">Operadora de Celular</Label>
                  <Input
                    id="cellOperator"
                    value={formData.cellOperator}
                    onChange={(e) => handleInputChange("cellOperator", e.target.value)}
                    placeholder="Ex: Vivo, Claro, Tim"
                  />
                </div>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    id="cellWifi"
                    checked={formData.cellWifi}
                    onCheckedChange={(checked) => handleInputChange("cellWifi", checked)}
                  />
                  <Label htmlFor="cellWifi" className="font-normal cursor-pointer">
                    WiFi Disponível
                  </Label>
                </div>
              </div>

              <div className="mt-8 space-y-4">
                <h3 className="font-semibold">Avaliação de Dificuldade</h3>
                <div>
                  <Label>Dificuldade Logística: {formData.logisticDifficulty}/10</Label>
                  <Slider
                    value={[formData.logisticDifficulty]}
                    onValueChange={(value) =>
                      handleInputChange("logisticDifficulty", value[0])
                    }
                    min={0}
                    max={10}
                    step={1}
                    className="mt-2"
                  />
                </div>
                <div>
                  <Label>Dificuldade de Montagem: {formData.assemblyDifficulty}/10</Label>
                  <Slider
                    value={[formData.assemblyDifficulty]}
                    onValueChange={(value) =>
                      handleInputChange("assemblyDifficulty", value[0])
                    }
                    min={0}
                    max={10}
                    step={1}
                    className="mt-2"
                  />
                </div>
              </div>
            </TabsContent>

            {/* Aba Observações */}
            <TabsContent value="observacoes" className="space-y-6 mt-6">
              <div>
                <Label htmlFor="observations">Observações Gerais</Label>
                <Textarea
                  id="observations"
                  value={formData.observations}
                  onChange={(e) => handleInputChange("observations", e.target.value)}
                  placeholder="Descreva as observações da visita..."
                  rows={4}
                />
              </div>

              <div>
                <Label htmlFor="clientActions">Ações do Cliente</Label>
                <Textarea
                  id="clientActions"
                  value={formData.clientActions}
                  onChange={(e) => handleInputChange("clientActions", e.target.value)}
                  placeholder="Descreva as ações que o cliente deve tomar..."
                  rows={4}
                />
              </div>

              <div>
                <Label htmlFor="supergasActions">Ações da Supergasbras</Label>
                <Textarea
                  id="supergasActions"
                  value={formData.supergasActions}
                  onChange={(e) => handleInputChange("supergasActions", e.target.value)}
                  placeholder="Descreva as ações que a Supergasbras deve tomar..."
                  rows={4}
                />
              </div>

              {/* Fotos */}
              <div className="mt-8">
                <h3 className="font-semibold text-lg mb-4">Fotos da Visita</h3>

                <Card>
                  <CardHeader>
                    <CardTitle>Adicionar Fotos</CardTitle>
                    <CardDescription>
                      Adicione fotos com observações/legendas
                    </CardDescription>
                  </CardHeader>
                  <CardContent className="space-y-4">
                    <div>
                      <Label htmlFor="photo-input">Selecionar Foto</Label>
                      <input
                        ref={fileInputRef}
                        id="photo-input"
                        type="file"
                        accept="image/*"
                        onChange={handleFileSelect}
                        className="hidden"
                      />
                      <Button
                        variant="outline"
                        onClick={() => fileInputRef.current?.click()}
                        className="mt-2 w-full"
                      >
                        <Upload className="w-4 h-4 mr-2" />
                        Escolher Foto
                      </Button>
                    </div>

                    <div>
                      <Label htmlFor="photo-observation">
                        Observação/Legenda da Foto
                      </Label>
                      <Textarea
                        id="photo-observation"
                        value={currentPhotoObservation}
                        onChange={(e) => setCurrentPhotoObservation(e.target.value)}
                        placeholder="Ex: Fachada do cliente, Local da central de gás, etc"
                        rows={2}
                      />
                    </div>
                  </CardContent>
                </Card>

                {/* Galeria de Fotos */}
                {photos.length > 0 && (
                  <div className="mt-6">
                    <h4 className="font-semibold mb-4">
                      Fotos Adicionadas ({photos.length})
                    </h4>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                      {photos.map((photo) => (
                        <div key={photo.id} className="relative group">
                          <div className="aspect-square bg-gray-100 rounded-lg overflow-hidden">
                            <img
                              src={photo.dataUrl}
                              alt="Foto"
                              className="w-full h-full object-cover"
                            />
                          </div>
                          {photo.observation && (
                            <p className="mt-2 text-sm text-gray-600 line-clamp-2">
                              {photo.observation}
                            </p>
                          )}
                          <Button
                            variant="destructive"
                            size="sm"
                            className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity"
                            onClick={() => handleDeletePhoto(photo.id)}
                          >
                            <X className="w-4 h-4" />
                          </Button>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {photos.length === 0 && (
                  <div className="flex flex-col items-center justify-center py-8 text-center mt-6">
                    <ImageIcon className="w-12 h-12 text-gray-300 mb-3" />
                    <p className="text-gray-500">Nenhuma foto adicionada ainda</p>
                  </div>
                )}
              </div>
            </TabsContent>
          </Tabs>
        </div>
      </div>
    </div>
  );
}
