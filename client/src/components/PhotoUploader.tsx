import { useState, useRef } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { trpc } from "@/lib/trpc";
import { toast } from "sonner";
import { Upload, X, Loader2, Image as ImageIcon } from "lucide-react";

interface Photo {
  id: number;
  photoUrl: string;
  photoKey: string;
  caption?: string | null;
  order: number | null;
  createdAt?: Date;
}

interface PhotoUploaderProps {
  reportId: number;
  photos: Photo[];
  onPhotosChange: (photos: Photo[]) => void;
}

export default function PhotoUploader({ reportId, photos, onPhotosChange }: PhotoUploaderProps) {
  const [isUploading, setIsUploading] = useState(false);
  const [caption, setCaption] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  const addPhotoMutation = trpc.photos.add.useMutation();
  const deletePhotoMutation = trpc.photos.delete.useMutation();

  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Validate file type
    if (!file.type.startsWith("image/")) {
      toast.error("Por favor, selecione uma imagem válida");
      return;
    }

    // Validate file size (max 5MB)
    if (file.size > 5 * 1024 * 1024) {
      toast.error("A imagem deve ter menos de 5MB");
      return;
    }

    setIsUploading(true);
    try {
      // In a real implementation, you would upload to S3 here
      // For now, we'll use a data URL
      const reader = new FileReader();
      reader.onload = async (event) => {
        const dataUrl = event.target?.result as string;

        // Add photo to database
        await addPhotoMutation.mutateAsync({
          reportId,
          photoUrl: dataUrl,
          photoKey: `report-${reportId}-${Date.now()}`,
          caption: caption || undefined,
          order: photos.length,
        });

        toast.success("Foto adicionada com sucesso!");
        setCaption("");
        if (fileInputRef.current) {
          fileInputRef.current.value = "";
        }

        // Refetch photos
        const photosResponse = await trpc.photos.getByReportId.useQuery({ reportId });
        if (photosResponse.data) {
          onPhotosChange(photosResponse.data);
        }
      };
      reader.readAsDataURL(file);
    } catch (error: any) {
      toast.error(error.message || "Erro ao fazer upload da foto");
    } finally {
      setIsUploading(false);
    }
  };

  const handleDeletePhoto = async (photoId: number) => {
    try {
      await deletePhotoMutation.mutateAsync({ id: photoId });
      toast.success("Foto removida com sucesso!");
      onPhotosChange(photos.filter((p) => p.id !== photoId));
    } catch (error: any) {
      toast.error(error.message || "Erro ao remover foto");
    }
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle>Fotos da Visita</CardTitle>
        <CardDescription>
          Adicione fotos para documentar a visita técnica
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        {/* Upload Area */}
        <div className="space-y-4">
          <div>
            <Label htmlFor="photo-input">Selecionar Foto</Label>
            <div className="mt-2 flex items-center gap-4">
              <input
                ref={fileInputRef}
                id="photo-input"
                type="file"
                accept="image/*"
                onChange={handleFileSelect}
                disabled={isUploading}
                className="hidden"
              />
              <Button
                variant="outline"
                onClick={() => fileInputRef.current?.click()}
                disabled={isUploading}
              >
                {isUploading ? (
                  <>
                    <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                    Enviando...
                  </>
                ) : (
                  <>
                    <Upload className="w-4 h-4 mr-2" />
                    Escolher Foto
                  </>
                )}
              </Button>
            </div>
          </div>

          <div>
            <Label htmlFor="caption">Descrição/Legenda (Opcional)</Label>
            <Textarea
              id="caption"
              value={caption}
              onChange={(e) => setCaption(e.target.value)}
              placeholder="Descreva o que está na foto (ex: Fachada do cliente, Local da central de gás)"
              rows={2}
              disabled={isUploading}
            />
          </div>
        </div>

        {/* Photos Grid */}
        {photos.length > 0 && (
          <div>
            <h3 className="font-semibold mb-4">Fotos Adicionadas ({photos.length})</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {photos.map((photo) => (
                <div key={photo.id} className="relative group">
                  <div className="aspect-square bg-gray-100 rounded-lg overflow-hidden">
                    <img
                      src={photo.photoUrl}
                      alt={photo.caption || "Foto"}
                      className="w-full h-full object-cover"
                    />
                  </div>
                  {photo.caption && (
                    <p className="mt-2 text-sm text-gray-600 line-clamp-2">
                      {photo.caption}
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
          <div className="flex flex-col items-center justify-center py-8 text-center">
            <ImageIcon className="w-12 h-12 text-gray-300 mb-3" />
            <p className="text-gray-500">Nenhuma foto adicionada ainda</p>
          </div>
        )}
      </CardContent>
    </Card>
  );
}
