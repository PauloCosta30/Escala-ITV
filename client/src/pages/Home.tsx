import { useAuth } from "@/_core/hooks/useAuth";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { useLocation } from "wouter";
import { getLoginUrl } from "@/const";
import { FileText, Plus, BarChart3, Clock } from "lucide-react";

export default function Home() {
  const { user, isAuthenticated } = useAuth();
  const [, navigate] = useLocation();

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100">
        <div className="max-w-6xl mx-auto px-4 py-20">
          <div className="text-center mb-12">
            <h1 className="text-4xl font-bold text-gray-900 mb-4">
              Sistema de Relatório de Visita Técnica
            </h1>
            <p className="text-xl text-gray-600 mb-8">
              Gerenciamento completo de relatórios de visita técnica para a Supergasbras
            </p>
            <Button size="lg" asChild>
              <a href={getLoginUrl()}>
                Entrar com Manus
              </a>
            </Button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <Card>
              <CardHeader>
                <FileText className="w-8 h-8 text-blue-600 mb-2" />
                <CardTitle>Formulário Completo</CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-gray-600">
                  Preencha todos os dados da visita técnica com campos estruturados e validados
                </p>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <Plus className="w-8 h-8 text-green-600 mb-2" />
                <CardTitle>Upload de Fotos</CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-gray-600">
                  Adicione múltiplas fotos com legendas descritivas para documentar a visita
                </p>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <BarChart3 className="w-8 h-8 text-purple-600 mb-2" />
                <CardTitle>Exportação em PDF</CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-gray-600">
                  Exporte relatórios completos em PDF com formatação profissional
                </p>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-6xl mx-auto px-4 py-12">
        <div className="mb-12">
          <h1 className="text-3xl font-bold text-gray-900 mb-2">
            Bem-vindo, {user?.name || "Usuário"}!
          </h1>
          <p className="text-gray-600">
            Gerencie seus relatórios de visita técnica
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-12">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <Plus className="w-5 h-5" />
                Novo Relatório
              </CardTitle>
              <CardDescription>
                Criar um novo relatório de visita técnica
              </CardDescription>
            </CardHeader>
            <CardContent>
              <p className="text-gray-600 mb-4">
                Inicie um novo relatório preenchendo os dados do projetista, cliente, central e outras informações relevantes.
              </p>
              <Button onClick={() => navigate("/reports/new")} className="w-full">
                Criar Novo Relatório
              </Button>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                <FileText className="w-5 h-5" />
                Meus Relatórios
              </CardTitle>
              <CardDescription>
                Visualizar e gerenciar seus relatórios
              </CardDescription>
            </CardHeader>
            <CardContent>
              <p className="text-gray-600 mb-4">
                Acesse todos os seus relatórios, edite-os, exporte em PDF ou delete conforme necessário.
              </p>
              <Button onClick={() => navigate("/reports")} variant="outline" className="w-full">
                Ver Meus Relatórios
              </Button>
            </CardContent>
          </Card>
        </div>

        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <Clock className="w-5 h-5" />
              Recursos Principais
            </CardTitle>
          </CardHeader>
          <CardContent>
            <ul className="space-y-3 text-gray-600">
              <li className="flex items-start gap-3">
                <span className="text-green-600 font-bold">✓</span>
                <span>Formulário completo com todos os campos do modelo de relatório</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="text-green-600 font-bold">✓</span>
                <span>Sistema de upload e visualização de múltiplas fotos com legendas</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="text-green-600 font-bold">✓</span>
                <span>Exportação de relatórios em PDF com formatação profissional</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="text-green-600 font-bold">✓</span>
                <span>Salvamento automático de rascunhos para evitar perda de dados</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="text-green-600 font-bold">✓</span>
                <span>Checklist interativo com questões Sim/Não</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="text-green-600 font-bold">✓</span>
                <span>Avaliação de dificuldade com escala de 0 a 10</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="text-green-600 font-bold">✓</span>
                <span>Notificações automáticas ao proprietário</span>
              </li>
              <li className="flex items-start gap-3">
                <span className="text-green-600 font-bold">✓</span>
                <span>Listagem com opções de visualizar, editar e deletar</span>
              </li>
            </ul>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
