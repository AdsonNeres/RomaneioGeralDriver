import React, { useState, useCallback } from 'react';
import { Upload, Download, FileSpreadsheet, AlertCircle, CheckCircle, Loader, X, RefreshCw, ArrowLeft, Truck, Package } from 'lucide-react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

interface ImportedData {
  servico: string;
  endereco: string;
}

interface ConsolidatedData {
  servicos: string[];
  endereco: string;
}

type ProcessingStatus = 'idle' | 'processing' | 'success' | 'error';
type AppMode = 'home' | 'vuupt' | 'tms';

const cleanAddress = (address: string): string => {
  // Lista de prefixos a serem removidos (case insensitive)
  const prefixesToRemove = [
    /^Condomínio\s+[^-]+-\s*/i,
    /^Condominio\s+[^-]+-\s*/i,
    /^Edifício\s+[^-]+-\s*/i,
    /^Edificio\s+[^-]+-\s*/i,
    /^Cond\.\s+[^-]+-\s*/i,
    /^Prédio\s+[^-]+-\s*/i,
    /^Predio\s+[^-]+-\s*/i
  ];

  let cleanedAddress = address.trim();
  
  // Remove cada prefixo se encontrado
  for (const prefix of prefixesToRemove) {
    cleanedAddress = cleanedAddress.replace(prefix, '');
  }
  
  return cleanedAddress.trim();
};

function App() {
  const [mode, setMode] = useState<AppMode>('home');
  const [file, setFile] = useState<File | null>(null);
  const [data, setData] = useState<ImportedData[]>([]);
  const [consolidatedData, setConsolidatedData] = useState<ConsolidatedData[]>([]);
  const [consolidateAddresses, setConsolidateAddresses] = useState(true);
  const [status, setStatus] = useState<ProcessingStatus>('idle');
  const [error, setError] = useState<string>('');
  const [isDragging, setIsDragging] = useState(false);
  const [showDriverModal, setShowDriverModal] = useState(false);
  const [driverName, setDriverName] = useState('');
  const [tempDriverName, setTempDriverName] = useState('');
  const [showConverterModal, setShowConverterModal] = useState(false);
  const [converterFile, setConverterFile] = useState<File | null>(null);
  const [converterStatus, setConverterStatus] = useState<ProcessingStatus>('idle');
  const [converterError, setConverterError] = useState<string>('');

  // Estados específicos para TMS
  const [tmsFile, setTmsFile] = useState<File | null>(null);
  const [tmsData, setTmsData] = useState<ImportedData[]>([]);
  const [tmsConsolidatedData, setTmsConsolidatedData] = useState<ConsolidatedData[]>([]);
  const [tmsConsolidateAddresses, setTmsConsolidateAddresses] = useState(true);
  const [tmsStatus, setTmsStatus] = useState<ProcessingStatus>('idle');
  const [tmsError, setTmsError] = useState<string>('');
  const [tmsIsDragging, setTmsIsDragging] = useState(false);
  const [showTmsDriverModal, setShowTmsDriverModal] = useState(false);
  const [tmsDriverName, setTmsDriverName] = useState('');
  const [tempTmsDriverName, setTempTmsDriverName] = useState('');

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = Array.from(e.dataTransfer.files);
    const xlsxFile = files.find(file => 
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.name.endsWith('.xlsx')
    );
    if (xlsxFile) {
      processFile(xlsxFile);
    }
  }, []);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      processFile(selectedFile);
    }
  }, []);

  // Handlers específicos para TMS
  const handleTmsDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setTmsIsDragging(true);
  }, []);

  const handleTmsDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setTmsIsDragging(false);
  }, []);

  const handleTmsDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setTmsIsDragging(false);
    const files = Array.from(e.dataTransfer.files);
    const xlsxFile = files.find(file => 
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.name.endsWith('.xlsx')
    );
    if (xlsxFile) {
      processTmsFile(xlsxFile);
    }
  }, []);

  const handleTmsFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      processTmsFile(selectedFile);
    }
  }, []);

  const processFile = useCallback((selectedFile: File) => {
    setFile(selectedFile);
    setStatus('processing');
    setError('');
    setData([]);
    setConsolidatedData([]);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target?.result, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Encontrar as colunas "Endereço" e "Serviço"
        // Extrair dados diretamente da coluna A e C a partir da linha 11 (índice 10)
        const extractedData: ImportedData[] = [];
        const startRow = 9; // Linha 11 (índice 10)
        
        for (let i = startRow; i < jsonData.length; i++) {
          const row = jsonData[i] as string[];
          if (!row) continue;
          
          // Coluna A (índice 0) = Serviço
          // Coluna C (índice 2) = Endereço
          const servico = row[0];
          const endereco = row[2];
          
          // Verificar se temos dados válidos (pular linhas vazias)
          if (servico && endereco) {
            extractedData.push({
              servico: String(servico).trim(),
              endereco: cleanAddress(String(endereco).trim())
            });
          }
        }

        if (extractedData.length === 0) {
          throw new Error('Nenhum dado válido encontrado nas colunas A e C a partir da linha 11');
        }

        setData(extractedData);
        
        // Sempre consolidar os dados para ter disponível
        const addressMap = new Map<string, string[]>();
        
        extractedData.forEach(item => {
          if (addressMap.has(item.endereco)) {
            addressMap.get(item.endereco)!.push(item.servico);
          } else {
            addressMap.set(item.endereco, [item.servico]);
          }
        });
        
        const consolidated: ConsolidatedData[] = Array.from(addressMap.entries()).map(([endereco, servicos]) => ({
          endereco,
          servicos
        }));
        
        setConsolidatedData(consolidated);
        
        setStatus('success');
        setShowDriverModal(true);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Erro ao processar arquivo');
        setStatus('error');
      }
    };

    reader.onerror = () => {
      setError('Erro ao ler o arquivo');
      setStatus('error');
    };

    reader.readAsBinaryString(selectedFile);
  }, []);

  const processTmsFile = useCallback((selectedFile: File) => {
    setTmsFile(selectedFile);
    setTmsStatus('processing');
    setTmsError('');
    setTmsData([]);
    setTmsConsolidatedData([]);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target?.result, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Processar dados TMS com colunas específicas
        const headers = jsonData[0] as string[];
        const extractedData: ImportedData[] = [];
        
        // Procurar pelas colunas específicas do TMS
        let numeroOrdemIndex = -1;
        let logradouroIndex = -1;
        let numeroIndex = -1;
        let bairroIndex = -1;
        let cidadeIndex = -1;
        let ufIndex = -1;
        let cepIndex = -1;
        
        headers.forEach((header, index) => {
          const headerLower = String(header).toLowerCase();
          if (headerLower.includes('numero_da_ordem') || headerLower.includes('ordem')) {
            numeroOrdemIndex = index;
          }
          if (headerLower.includes('logradouro_destinatario') || headerLower.includes('logradouro')) {
            logradouroIndex = index;
          }
          if (headerLower.includes('numero_do_destinatario') || headerLower.includes('numero_destinatario')) {
            numeroIndex = index;
          }
          if (headerLower.includes('bairro_destinatario') || headerLower.includes('bairro')) {
            bairroIndex = index;
          }
          if (headerLower.includes('cidade_destinatario') || headerLower.includes('cidade')) {
            cidadeIndex = index;
          }
          if (headerLower.includes('uf_destinatario') || headerLower.includes('uf')) {
            ufIndex = index;
          }
          if (headerLower.includes('cep_destinatario') || headerLower.includes('cep')) {
            cepIndex = index;
          }
        });
        
        // Verificar se encontrou as colunas obrigatórias
        if (numeroOrdemIndex === -1) {
          throw new Error('Coluna "Numero_da_Ordem" não encontrada');
        }
        if (logradouroIndex === -1) {
          throw new Error('Coluna "Logradouro_Destinatario" não encontrada');
        }
        
        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i] as string[];
          if (!row) continue;
          
          const numeroOrdem = row[numeroOrdemIndex];
          const logradouro = row[logradouroIndex];
          const numero = row[numeroIndex] || '';
          const bairro = row[bairroIndex] || '';
          const cidade = row[cidadeIndex] || '';
          const uf = row[ufIndex] || '';
          const cep = row[cepIndex] || '';
          
          // Montar o endereço no formato desejado
          // "R. Barra do Turvo, 49 - Cordovil, Rio de Janeiro - RJ, 21010-220, Brasil"
          let enderecoCompleto = '';
          
          if (logradouro) {
            enderecoCompleto += String(logradouro).trim();
          }
          
          if (numero) {
            enderecoCompleto += `, ${String(numero).trim()}`;
          }
          
          if (bairro) {
            enderecoCompleto += ` - ${String(bairro).trim()}`;
          }
          
          if (cidade) {
            enderecoCompleto += `, ${String(cidade).trim()}`;
          }
          
          if (uf) {
            enderecoCompleto += ` - ${String(uf).trim()}`;
          }
          
          if (cep) {
            enderecoCompleto += `, ${String(cep).trim()}`;
          }
          
          // Adicionar "Brasil" no final
          if (enderecoCompleto) {
            enderecoCompleto += ', Brasil';
          }
          
          if (numeroOrdem && enderecoCompleto) {
            extractedData.push({
              servico: String(numeroOrdem).trim(),
              endereco: cleanAddress(enderecoCompleto.trim())
            });
          }
        }

        if (extractedData.length === 0) {
          throw new Error('Nenhum dado válido encontrado. Verifique se o arquivo contém as colunas necessárias.');
        }

        setTmsData(extractedData);
        
        // Processar dados para adicionar sufixos em ordens duplicadas
        const processedData: ImportedData[] = [];
        const orderCountMap = new Map<string, number>();
        
        // Primeiro, contar quantas vezes cada ordem aparece
        extractedData.forEach(item => {
          const count = orderCountMap.get(item.servico) || 0;
          orderCountMap.set(item.servico, count + 1);
        });
        
        // Depois, processar os dados adicionando sufixos quando necessário
        const orderSuffixMap = new Map<string, number>();
        
        extractedData.forEach(item => {
          const originalOrder = item.servico;
          const totalCount = orderCountMap.get(originalOrder) || 1;
          
          if (totalCount > 1) {
            // Se há múltiplas ocorrências, adicionar sufixo
            const currentSuffix = orderSuffixMap.get(originalOrder) || 0;
            const suffixLetter = String.fromCharCode(65 + currentSuffix); // A, B, C, etc.
            const newOrder = `${originalOrder}${suffixLetter}`;
            
            processedData.push({
              servico: newOrder,
              endereco: item.endereco
            });
            
            orderSuffixMap.set(originalOrder, currentSuffix + 1);
          } else {
            // Se há apenas uma ocorrência, manter como está
            processedData.push(item);
          }
        });
        
        // Atualizar os dados processados
        setTmsData(processedData);
        
        // Consolidar dados se necessário (para a opção de consolidação)
        const addressMap = new Map<string, string[]>();
        
        processedData.forEach(item => {
          if (addressMap.has(item.endereco)) {
            addressMap.get(item.endereco)!.push(item.servico);
          } else {
            addressMap.set(item.endereco, [item.servico]);
          }
        });
        
        const consolidated: ConsolidatedData[] = Array.from(addressMap.entries()).map(([endereco, servicos]) => ({
          endereco,
          servicos
        }));
        
        setTmsConsolidatedData(consolidated);
        
        setTmsStatus('success');
        setShowTmsDriverModal(true);
      } catch (err) {
        setTmsError(err instanceof Error ? err.message : 'Erro ao processar arquivo');
        setTmsStatus('error');
      }
    };

    reader.onerror = () => {
      setTmsError('Erro ao ler o arquivo');
      setTmsStatus('error');
    };

    reader.readAsBinaryString(selectedFile);
  }, []);

  const handleDriverNameSubmit = useCallback(() => {
    if (tempDriverName.trim()) {
      setDriverName(tempDriverName.trim());
      setShowDriverModal(false);
      setTempDriverName('');
    }
  }, [tempDriverName]);

  const handleDriverModalClose = useCallback(() => {
    setShowDriverModal(false);
    setTempDriverName('');
  }, []);

  const handleTmsDriverNameSubmit = useCallback(() => {
    if (tempTmsDriverName.trim()) {
      setTmsDriverName(tempTmsDriverName.trim());
      setShowTmsDriverModal(false);
      setTempTmsDriverName('');
    }
  }, [tempTmsDriverName]);

  const handleTmsDriverModalClose = useCallback(() => {
    setShowTmsDriverModal(false);
    setTempTmsDriverName('');
  }, []);

  const handleConverterClick = useCallback(() => {
    setShowConverterModal(true);
    setConverterFile(null);
    setConverterStatus('idle');
    setConverterError('');
  }, []);

  const handleConverterFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) {
      setConverterFile(selectedFile);
    }
  }, []);

  const handleConvertFile = useCallback(() => {
    if (!converterFile) return;

    setConverterStatus('processing');
    setConverterError('');

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        // Ler o arquivo XLS
        const workbook = XLSX.read(e.target?.result, { type: 'binary' });
        
        // Consolidar todas as sheets em uma única
        const consolidatedData: any[][] = [];
        let isFirstSheet = true;
        
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          if (isFirstSheet) {
            // Para a primeira sheet, adiciona todos os dados
            consolidatedData.push(...sheetData);
            isFirstSheet = false;
          } else {
            // Para as outras sheets, pula a primeira linha (cabeçalho) se existir
            const dataWithoutHeader = sheetData.length > 0 ? sheetData.slice(1) : [];
            consolidatedData.push(...dataWithoutHeader);
          }
        });
        
        // Criar nova planilha com dados convertidos e fazer download imediatamente
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.aoa_to_sheet(consolidatedData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Dados Consolidados');
        
        // Converter para XLSX e fazer download
        const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // Nome do arquivo com timestamp
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
        const fileName = `Arquivo_Convertido_${timestamp}.xlsx`;
        
        // Fazer download
        try {
          saveAs(blob, fileName);
          console.log('Download iniciado:', fileName);
        } catch (saveError) {
          console.error('Erro com file-saver, tentando método alternativo:', saveError);
          
          // Método alternativo de download
          const url = window.URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = fileName;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          window.URL.revokeObjectURL(url);
          console.log('Download alternativo iniciado:', fileName);
        }
        
        setConverterStatus('success');
        console.log('Conversão concluída:', consolidatedData.length, 'linhas processadas');
        
      } catch (err) {
        setConverterError(err instanceof Error ? err.message : 'Erro ao converter arquivo');
        setConverterStatus('error');
      }
    };

    reader.onerror = () => {
      setConverterError('Erro ao ler o arquivo');
      setConverterStatus('error');
    };

    reader.readAsBinaryString(converterFile);
  }, [converterFile]);

  const handleCloseConverterModal = useCallback(() => {
    setShowConverterModal(false);
    setConverterFile(null);
    setConverterStatus('idle');
    setConverterError('');
  }, []);

  const generateNewSpreadsheet = useCallback(() => {
    if (data.length === 0 || !driverName) return;

    let newData: (string | number)[][];
    
    if (consolidateAddresses) {
      // Criar nova planilha consolidada
      newData = [
        ...consolidatedData.map(item => [item.servicos.join(', '), item.endereco])
      ];
    } else {
      // Criar nova planilha com Serviço na coluna A e Endereço na coluna B
      newData = [
        ...data.map(item => [item.servico, item.endereco])
      ];
    }

    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(newData);
    
    // Ajustar largura das colunas
    newWorksheet['!cols'] = [
      { width: consolidateAddresses ? 50 : 30 }, // Coluna A - Serviço(s)
      { width: 40 }  // Coluna B - Endereço
    ];

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Dados Processados');
    
    // Gerar arquivo e fazer download
    const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Romaneio_${driverName}.xlsx`);
  }, [data, consolidatedData, consolidateAddresses, driverName]);

  const generateTmsSpreadsheet = useCallback(() => {
    if (tmsData.length === 0 || !tmsDriverName) return;

    let newData: (string | number)[][];
    
    if (tmsConsolidateAddresses) {
      // Criar nova planilha consolidada
      newData = [
        ...tmsConsolidatedData.map(item => [item.servicos.join(', '), item.endereco])
      ];
    } else {
      // Criar nova planilha com Serviço na coluna A e Endereço na coluna B
      newData = [
        ...tmsData.map(item => [item.servico, item.endereco])
      ];
    }

    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(newData);
    
    // Ajustar largura das colunas
    newWorksheet['!cols'] = [
      { width: tmsConsolidateAddresses ? 50 : 30 }, // Coluna A - Serviço(s)
      { width: 40 }  // Coluna B - Endereço
    ];

    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Dados Processados TMS');
    
    // Gerar arquivo e fazer download
    const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Romaneio_TMS_${tmsDriverName}.xlsx`);
  }, [tmsData, tmsConsolidatedData, tmsConsolidateAddresses, tmsDriverName]);

  const resetState = useCallback(() => {
    setFile(null);
    setData([]);
    setConsolidatedData([]);
    setStatus('idle');
    setError('');
    setDriverName('');
    setShowDriverModal(false);
    setTempDriverName('');
  }, []);

  const resetTmsState = useCallback(() => {
    setTmsFile(null);
    setTmsData([]);
    setTmsConsolidatedData([]);
    setTmsStatus('idle');
    setTmsError('');
    setTmsDriverName('');
    setShowTmsDriverModal(false);
    setTempTmsDriverName('');
  }, []);

  const goHome = useCallback(() => {
    setMode('home');
    resetState();
    resetTmsState();
  }, [resetState, resetTmsState]);

  // Tela inicial com seleção de modo
  if (mode === 'home') {
    return (
      <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 flex items-center justify-center p-4">
        <div className="max-w-4xl w-full">
          {/* Header */}
          <div className="text-center mb-12">
            <div className="flex items-center justify-center mb-6">
              <Truck className="h-16 w-16 text-blue-600 mr-4" />
              <h1 className="text-5xl font-bold text-gray-800">Sistema de Romaneios</h1>
            </div>
            <p className="text-xl text-gray-600">
              Escolha o tipo de romaneio que deseja processar
            </p>
          </div>

          {/* Botões de seleção */}
          <div className="grid md:grid-cols-2 gap-8">
            {/* Romaneio Vuupt */}
            <div 
              onClick={() => setMode('vuupt')}
              className="group bg-white rounded-3xl shadow-xl p-8 cursor-pointer transform transition-all duration-300 hover:scale-105 hover:shadow-2xl border-2 border-transparent hover:border-blue-200"
            >
              <div className="text-center">
                <div className="bg-gradient-to-br from-blue-500 to-blue-600 rounded-2xl p-6 mb-6 inline-block group-hover:from-blue-600 group-hover:to-blue-700 transition-all duration-300">
                  <Package className="h-12 w-12 text-white" />
                </div>
                <h2 className="text-3xl font-bold text-gray-800 mb-20">Romaneio Vuupt</h2>
                
                <div className="bg-blue-50 rounded-xl p-4 text-left"> 
                  <p> Para fazer o romaneio Vuupt, precisa converter XLS no botão roxo</p>
                </div>
              </div>
            </div>

            {/* Romaneio TMS */}
            <div 
              onClick={() => setMode('tms')}
              className="group bg-white rounded-3xl shadow-xl p-8 cursor-pointer transform transition-all duration-300 hover:scale-105 hover:shadow-2xl border-2 border-transparent hover:border-green-200"
            >
              <div className="text-center">
                <div className="bg-gradient-to-br from-green-500 to-green-600 rounded-2xl p-6 mb-6 inline-block group-hover:from-green-600 group-hover:to-green-700 transition-all duration-300">
                  <FileSpreadsheet className="h-12 w-12 text-white" />
                </div>
                <h2 className="text-3xl font-bold text-gray-800 mb-20">Romaneio TMS</h2>
                <div className="bg-green-50 rounded-xl p-4 text-left">
                  <p className="text-green-700">
                    Para o romaneio TMS não precisa converter XLS, só baixar e importar aqui.
                  </p>
                </div>
              </div>
            </div>
          </div>

          {/* Footer */}
          <div className="text-center mt-12">
            <p className="text-gray-500">
              Selecione uma opção acima para começar
            </p>
          </div>
        </div>
      </div>
    );
  }

  // Tela do Romaneio TMS
  if (mode === 'tms') {
    return (
      <div className="min-h-screen bg-gradient-to-br from-green-50 to-emerald-50 p-4">
        <div className="max-w-6xl mx-auto">
          {/* Header com botão voltar */}
          <div className="flex items-center mb-8">
            <button
              onClick={goHome}
              className="flex items-center px-4 py-2 bg-white text-gray-700 rounded-xl font-semibold hover:bg-gray-50 transition-all duration-200 shadow-md mr-4"
            >
              <ArrowLeft className="h-4 w-4 mr-2" />
              Voltar
            </button>
            <div className="flex items-center">
              <FileSpreadsheet className="h-10 w-10 text-green-600 mr-3" />
              <h1 className="text-3xl font-bold text-gray-800">Romaneio TMS</h1>
            </div>
          </div>

          {/* Upload Area */}
          <div className="bg-white rounded-2xl shadow-xl p-8 mb-8">
            {/* Toggle para consolidar endereços */}
            <div className="mb-6 flex items-center justify-center">
              <label className="flex items-center cursor-pointer">
                <div className="relative">
                  <input
                    type="checkbox"
                    checked={tmsConsolidateAddresses}
                    onChange={(e) => setTmsConsolidateAddresses(e.target.checked)}
                    className="sr-only"
                  />
                  <div className={`block w-14 h-8 rounded-full transition-colors duration-200 ${
                    tmsConsolidateAddresses ? 'bg-green-600' : 'bg-gray-300'
                  }`}></div>
                  <div className={`absolute left-1 top-1 bg-white w-6 h-6 rounded-full transition-transform duration-200 ${
                    tmsConsolidateAddresses ? 'transform translate-x-6' : ''
                  }`}></div>
                </div>
                <span className="ml-3 text-gray-700 font-medium">
                  Consolidar endereços repetidos
                </span>
              </label>
            </div>
            
            <div
              className={`relative border-2 border-dashed rounded-xl p-12 text-center transition-all duration-300 ${
                tmsIsDragging
                  ? 'border-green-400 bg-green-50 scale-105'
                  : tmsStatus === 'success'
                  ? 'border-green-400 bg-green-50'
                  : tmsStatus === 'error'
                  ? 'border-red-400 bg-red-50'
                  : 'border-gray-300 bg-gray-50 hover:border-green-400 hover:bg-green-50'
              }`}
              onDragOver={handleTmsDragOver}
              onDragLeave={handleTmsDragLeave}
              onDrop={handleTmsDrop}
            >
              <input
                type="file"
                accept=".xlsx"
                onChange={handleTmsFileSelect}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              
              <div className="flex flex-col items-center space-y-4">
                {tmsStatus === 'processing' ? (
                  <>
                    <Loader className="h-16 w-16 text-green-600 animate-spin" />
                    <p className="text-xl font-semibold text-green-600">Processando arquivo...</p>
                  </>
                ) : tmsStatus === 'success' ? (
                  <>
                    <CheckCircle className="h-16 w-16 text-green-600" />
                    <p className="text-xl font-semibold text-green-600">
                      Arquivo processado com sucesso!
                    </p>
                    <p className="text-gray-600">
                      {tmsConsolidateAddresses 
                        ? `${tmsConsolidatedData.length} endereços únicos (${tmsData.length} registros originais)`
                        : `${tmsData.length} registros encontrados`
                      }
                    </p>
                  </>
                ) : tmsStatus === 'error' ? (
                  <>
                    <AlertCircle className="h-16 w-16 text-red-600" />
                    <p className="text-xl font-semibold text-red-600">Erro no processamento</p>
                    <p className="text-red-500">{tmsError}</p>
                  </>
                ) : (
                  <>
                    <Upload className="h-16 w-16 text-gray-400" />
                    <p className="text-xl font-semibold text-gray-700">
                      Arraste um arquivo XLSX aqui ou clique para selecionar
                    </p>
                    <p className="text-gray-500">
                      O sistema extrairá automaticamente os dados de Serviços / Endereços
                      {tmsConsolidateAddresses && (
                        <><br />Endereços repetidos serão consolidados em uma única linha</>
                      )}
                    </p>
                  </>
                )}
              </div>
            </div>

            {tmsFile && (
              <div className="mt-6 p-4 bg-green-50 rounded-lg">
                <p className="text-green-800">
                  <strong>Arquivo selecionado:</strong> {tmsFile.name}
                </p>
              </div>
            )}
          </div>

          {/* Action Buttons */}
          {(tmsStatus === 'success' || tmsStatus === 'error') && !showTmsDriverModal && (
            <div className="flex justify-center space-x-4 mb-8">
              {tmsStatus === 'success' && tmsDriverName && (
                <button
                  onClick={generateTmsSpreadsheet}
                  className="flex items-center px-8 py-3 bg-green-600 text-white rounded-xl font-semibold hover:bg-green-700 transition-all duration-200 transform hover:scale-105 shadow-lg"
                >
                  <Download className="h-5 w-5 mr-2" />
                  Gerar Romaneio TMS - {tmsDriverName}
                </button>
              )}
              <button
                onClick={resetTmsState}
                className="flex items-center px-8 py-3 bg-gray-600 text-white rounded-xl font-semibold hover:bg-gray-700 transition-all duration-200 transform hover:scale-105 shadow-lg"
              >
                Importar Novo Arquivo
              </button>
            </div>
          )}

          {/* Modal para Nome do Motorista TMS */}
          {showTmsDriverModal && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
              <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-8 transform transition-all duration-300">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="text-2xl font-bold text-gray-800">Nome do Motorista</h3>
                  <button
                    onClick={handleTmsDriverModalClose}
                    className="text-gray-400 hover:text-gray-600 transition-colors"
                  >
                    <X className="h-6 w-6" />
                  </button>
                </div>
                
                <div className="mb-6">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Digite o nome do motorista:
                  </label>
                  <input
                    type="text"
                    value={tempTmsDriverName}
                    onChange={(e) => setTempTmsDriverName(e.target.value)}
                    onKeyPress={(e) => e.key === 'Enter' && handleTmsDriverNameSubmit()}
                    placeholder="Ex: João"
                    className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-green-500 focus:border-transparent outline-none transition-all duration-200"
                    autoFocus
                  />
                </div>
                
                <div className="flex space-x-3">
                  <button
                    onClick={handleTmsDriverModalClose}
                    className="flex-1 px-4 py-3 bg-gray-200 text-gray-800 rounded-xl font-semibold hover:bg-gray-300 transition-all duration-200"
                  >
                    Cancelar
                  </button>
                  <button
                    onClick={handleTmsDriverNameSubmit}
                    disabled={!tempTmsDriverName.trim()}
                    className="flex-1 px-4 py-3 bg-green-600 text-white rounded-xl font-semibold hover:bg-green-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-all duration-200"
                  >
                    Confirmar
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Data Preview */}
          {tmsData.length > 0 && !showTmsDriverModal && (
            <div className="bg-white rounded-2xl shadow-xl overflow-hidden">
              <div className="p-6 bg-gradient-to-r from-green-600 to-green-700">
                <h2 className="text-2xl font-bold text-white">
                  Dados Processados TMS
                </h2>
              </div>
              
              <div className="p-8 text-center">
                <p className="text-gray-600 text-lg">
                  {tmsConsolidateAddresses 
                    ? `${tmsConsolidatedData.length} endereços únicos consolidados de ${tmsData.length} registros`
                    : `${tmsData.length} registros processados`
                  }
                </p>
                {tmsDriverName && (
                  <p className="text-green-600 font-semibold mt-2">
                    Motorista: {tmsDriverName}
                  </p>
                )}
                {tmsDriverName && (
                  <p className="text-gray-500 mt-2">
                    Clique em "Gerar Romaneio TMS" para fazer o download
                  </p>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    );
  }

  // Tela do Romaneio Vuupt (sistema atual)
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 p-4">
      <div className="max-w-6xl mx-auto">
        {/* Header com botão voltar */}
        <div className="relative mb-8">
          <div className="flex items-center">
            <button
              onClick={goHome}
              className="flex items-center px-4 py-2 bg-white text-gray-700 rounded-xl font-semibold hover:bg-gray-50 transition-all duration-200 shadow-md mr-4"
            >
              <ArrowLeft className="h-4 w-4 mr-2" />
              Voltar
            </button>
            <div className="flex items-center">
              <Package className="h-10 w-10 text-blue-600 mr-3" />
              <h1 className="text-3xl font-bold text-gray-800">Romaneio Vuupt</h1>
            </div>
          </div>
          
          {/* Botão Conversor no canto superior direito */}
          <div className="absolute top-0 right-0">
            <button
              onClick={handleConverterClick}
              className="flex items-center px-4 py-2 bg-purple-600 text-white rounded-xl font-semibold hover:bg-purple-700 transition-all duration-200 transform hover:scale-105 shadow-lg"
              title="Converter XLS para XLSX"
            >
              <RefreshCw className="h-4 w-4 mr-2" />
              Converter XLS
            </button>
          </div>
        </div>

        {/* Upload Area */}
        <div className="bg-white rounded-2xl shadow-xl p-8 mb-8">
          {/* Toggle para consolidar endereços */}
          <div className="mb-6 flex items-center justify-center">
            <label className="flex items-center cursor-pointer">
              <div className="relative">
                <input
                  type="checkbox"
                  checked={consolidateAddresses}
                  onChange={(e) => setConsolidateAddresses(e.target.checked)}
                  className="sr-only"
                />
                <div className={`block w-14 h-8 rounded-full transition-colors duration-200 ${
                  consolidateAddresses ? 'bg-blue-600' : 'bg-gray-300'
                }`}></div>
                <div className={`absolute left-1 top-1 bg-white w-6 h-6 rounded-full transition-transform duration-200 ${
                  consolidateAddresses ? 'transform translate-x-6' : ''
                }`}></div>
              </div>
              <span className="ml-3 text-gray-700 font-medium">
                Consolidar endereços repetidos
              </span>
            </label>
          </div>
          
          <div
            className={`relative border-2 border-dashed rounded-xl p-12 text-center transition-all duration-300 ${
              isDragging
                ? 'border-blue-400 bg-blue-50 scale-105'
                : status === 'success'
                ? 'border-green-400 bg-green-50'
                : status === 'error'
                ? 'border-red-400 bg-red-50'
                : 'border-gray-300 bg-gray-50 hover:border-blue-400 hover:bg-blue-50'
            }`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <input
              type="file"
              accept=".xlsx"
              onChange={handleFileSelect}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            />
            
            <div className="flex flex-col items-center space-y-4">
              {status === 'processing' ? (
                <>
                  <Loader className="h-16 w-16 text-blue-600 animate-spin" />
                  <p className="text-xl font-semibold text-blue-600">Processando arquivo...</p>
                </>
              ) : status === 'success' ? (
                <>
                  <CheckCircle className="h-16 w-16 text-green-600" />
                  <p className="text-xl font-semibold text-green-600">
                    Arquivo processado com sucesso!
                  </p>
                  <p className="text-gray-600">
                    {consolidateAddresses 
                      ? `${consolidatedData.length} endereços únicos (${data.length} registros originais)`
                      : `${data.length} registros encontrados`
                    }
                  </p>
                </>
              ) : status === 'error' ? (
                <>
                  <AlertCircle className="h-16 w-16 text-red-600" />
                  <p className="text-xl font-semibold text-red-600">Erro no processamento</p>
                  <p className="text-red-500">{error}</p>
                </>
              ) : (
                <>
                  <Upload className="h-16 w-16 text-gray-400" />
                  <p className="text-xl font-semibold text-gray-700">
                    Arraste um arquivo XLSX aqui ou clique para selecionar
                  </p>
                  <p className="text-gray-500">
                    O sistema extrairá automaticamente os dados de Serviços / Endereços
                    {consolidateAddresses && (
                      <><br />Endereços repetidos serão consolidados em uma única linha</>
                    )}
                  </p>
                </>
              )}
            </div>
          </div>

          {file && (
            <div className="mt-6 p-4 bg-blue-50 rounded-lg">
              <p className="text-blue-800">
                <strong>Arquivo selecionado:</strong> {file.name}
              </p>
            </div>
          )}
        </div>

        {/* Action Buttons */}
        {(status === 'success' || status === 'error') && !showDriverModal && (
          <div className="flex justify-center space-x-4 mb-8">
            {status === 'success' && driverName && (
              <button
                onClick={generateNewSpreadsheet}
                className="flex items-center px-8 py-3 bg-green-600 text-white rounded-xl font-semibold hover:bg-green-700 transition-all duration-200 transform hover:scale-105 shadow-lg"
              >
                <Download className="h-5 w-5 mr-2" />
                Gerar Romaneio - {driverName}
              </button>
            )}
            <button
              onClick={resetState}
              className="flex items-center px-8 py-3 bg-gray-600 text-white rounded-xl font-semibold hover:bg-gray-700 transition-all duration-200 transform hover:scale-105 shadow-lg"
            >
              Importar Novo Arquivo
            </button>
          </div>
        )}

        {/* Modal para Nome do Motorista */}
        {showDriverModal && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-8 transform transition-all duration-300">
              <div className="flex items-center justify-between mb-6">
                <h3 className="text-2xl font-bold text-gray-800">Nome do Motorista</h3>
                <button
                  onClick={handleDriverModalClose}
                  className="text-gray-400 hover:text-gray-600 transition-colors"
                >
                  <X className="h-6 w-6" />
                </button>
              </div>
              
              <div className="mb-6">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Digite o nome do motorista:
                </label>
                <input
                  type="text"
                  value={tempDriverName}
                  onChange={(e) => setTempDriverName(e.target.value)}
                  onKeyPress={(e) => e.key === 'Enter' && handleDriverNameSubmit()}
                  placeholder="Ex: Adson"
                  className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all duration-200"
                  autoFocus
                />
              </div>
              
              <div className="flex space-x-3">
                <button
                  onClick={handleDriverModalClose}
                  className="flex-1 px-4 py-3 bg-gray-200 text-gray-800 rounded-xl font-semibold hover:bg-gray-300 transition-all duration-200"
                >
                  Cancelar
                </button>
                <button
                  onClick={handleDriverNameSubmit}
                  disabled={!tempDriverName.trim()}
                  className="flex-1 px-4 py-3 bg-blue-600 text-white rounded-xl font-semibold hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-all duration-200"
                >
                  Confirmar
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal do Conversor XLS para XLSX */}
        {showConverterModal && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full p-8 transform transition-all duration-300">
              <div className="flex items-center justify-between mb-6">
                <h3 className="text-2xl font-bold text-gray-800">Converter XLS para XLSX</h3>
                <button
                  onClick={handleCloseConverterModal}
                  className="text-gray-400 hover:text-gray-600 transition-colors"
                >
                  <X className="h-6 w-6" />
                </button>
              </div>
              
              <div className="mb-6">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Selecione um arquivo XLS:
                </label>
                <input
                  type="file"
                  accept=".xls"
                  onChange={handleConverterFileSelect}
                  className="w-full px-4 py-3 border border-gray-300 rounded-xl focus:ring-2 focus:ring-purple-500 focus:border-transparent outline-none transition-all duration-200"
                />
                
                {converterFile && (
                  <div className="mt-3 p-3 bg-purple-50 rounded-lg">
                    <p className="text-purple-800 text-sm">
                      <strong>Arquivo selecionado:</strong> {converterFile.name}
                    </p>
                  </div>
                )}
                
                {converterStatus === 'processing' && (
                  <div className="mt-4 flex items-center justify-center text-purple-600">
                    <Loader className="h-5 w-5 animate-spin mr-2" />
                    Convertendo arquivo...
                  </div>
                )}
                
                {converterStatus === 'success' && (
                  <div className="mt-4 flex items-center justify-center text-green-600">
                    <CheckCircle className="h-5 w-5 mr-2" />
                    Arquivo convertido e baixado com sucesso!
                  </div>
                )}
                
                {converterStatus === 'error' && (
                  <div className="mt-4 p-3 bg-red-50 rounded-lg">
                    <div className="flex items-center text-red-600">
                      <AlertCircle className="h-5 w-5 mr-2" />
                      <span className="text-sm">{converterError}</span>
                    </div>
                  </div>
                )}
              </div>
              
              <div className="flex space-x-3">
                <button
                  onClick={handleCloseConverterModal}
                  className="flex-1 px-4 py-3 bg-gray-200 text-gray-800 rounded-xl font-semibold hover:bg-gray-300 transition-all duration-200"
                >
                  Cancelar
                </button>
                <button
                  onClick={handleConvertFile}
                  disabled={!converterFile || converterStatus === 'processing'}
                  className="flex-1 px-4 py-3 bg-purple-600 text-white rounded-xl font-semibold hover:bg-purple-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-all duration-200"
                >
                  {converterStatus === 'processing' ? 'Convertendo...' : 'Converter'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Data Preview */}
        {data.length > 0 && !showDriverModal && (
          <div className="bg-white rounded-2xl shadow-xl overflow-hidden">
            <div className="p-6 bg-gradient-to-r from-blue-600 to-blue-700">
              <h2 className="text-2xl font-bold text-white">
                Dados Processados
              </h2>
            </div>
            
            <div className="p-8 text-center">
              <p className="text-gray-600 text-lg">
                {consolidateAddresses 
                  ? `${consolidatedData.length} endereços únicos consolidados de ${data.length} registros`
                  : `${data.length} registros processados`
                }
              </p>
              {driverName && (
                <p className="text-blue-600 font-semibold mt-2">
                  Motorista: {driverName}
                </p>
              )}
              {driverName && (
                <p className="text-gray-500 mt-2">
                  Clique em "Gerar Romaneio" para fazer o download
                </p>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;