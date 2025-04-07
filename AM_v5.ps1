#FUNÇÃO QUE ESTABELCE UM TEMPO DE ESPERA PARA O ARQUIVO SER CRIADO APÓS IDENTIFICADO
function AguardarArquivo {
    param (
        [string]$caminhoArquivo,
        [int]$timeoutSegundos = 5
    )
    
    $tempoInicial = Get-Date
    while ($true) {
        # Verifica existência E tenta abrir o arquivo
        if (Test-Path -Path $caminhoArquivo -ErrorAction SilentlyContinue) {
            try {
                $fileStream = [System.IO.File]::Open($caminhoArquivo, 'Open', 'Read', 'None')
                $fileStream.Close()
                return $true
            } catch {
                # Arquivo existe mas não está acessível
                if (((Get-Date) - $tempoInicial).TotalSeconds -gt $timeoutSegundos) {
                    return $false
                }
                Start-Sleep -Milliseconds 100
            }
        } else {
            # Arquivo não existe
            if (((Get-Date) - $tempoInicial).TotalSeconds -gt $timeoutSegundos) {
                return $false
            }
            Start-Sleep -Milliseconds 100
        }
    }
}


################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################

# FUNÇÃO PARA MOVER O ARQUIVO DE UM DESTINO PARA O OUTRO
function MoverArquivo {
    param (
        [string]$sourcePath,
        [string]$destinoPasta
    )

    $nomeArquivo = [System.IO.Path]::GetFileName($sourcePath)
    $destinoCompleto = Join-Path -Path $destinoPasta -ChildPath $nomeArquivo

    try {
        Move-Item -Path $sourcePath -Destination $destinoCompleto -Force
        Write-Host "[INFO] Arquivo $nomeArquivo movido para $destinoPasta com sucesso." #APENAS VERIFICAÇÃO
        return $destinoCompleto
    } catch {
        Write-Host "[ERRO] Erro ao mover o arquivo: $_" #APENAS VERIFICAÇÃO / #TRATAMENTO DE ERRO
        return $null
    }
}


################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################

# FUNÇÃO QUE LÊ UM ARQUIVO TXT E TRATA AS INFORMAÇÕES
function TratarArquivoTXT {
    param (
        [string]$caminhoArquivo,
        [bool]$modoRapido = $false
    )

    try {
        #LÊ TODAS AS LINHAS DO ARQUIVO, REMOVE VAZIAS E ASTERISCOS
        $lines = Get-Content $caminhoArquivo | Where-Object { $_.Trim() -ne "" } | ForEach-Object { $_ -replace "\*", "" }

        if ($lines.Count -lt 3) {
            Write-Host "[AVISO] Arquivo com formato inválido" #APENAS VERIFICAÇÃO
            return $null
        }

        #PRIMEIRA LINHA: NÚMERO DO ITEM
        $global:numero_item = $lines[0].Trim()

        #CRIA LISTA PARA ARMAZENAR OS DADOS PROCESSADOS
        $valores_list = [System.Collections.Generic.List[PSCustomObject]]::new()

        #PROCESSA DA LINHA 3 EM DIANTE (ÍNDICE 2)
        for ($i = 2; $i -lt $lines.Count; $i++) {
            #DIVIDE USANDO QUALQUER ENTIDADE DE ESPAÇOS COMO SEPARADOR
            $cols = ($lines[$i].Trim() -split '\s+')

            if ($cols.Count -eq 4 -and $cols -notcontains "") {
                try {
                    $nom_val  = [math]::Abs([double]$cols[0])
                    $u_tol    = [math]::Abs([double]$cols[1])
                    $l_tol    = [math]::Abs([double]$cols[2])
                    $act_val  = [math]::Abs([double]$cols[3])

                    $culture = [System.Globalization.CultureInfo]::InvariantCulture

                    $max = $nom_val + $u_tol
                    $min = $nom_val - $l_tol

                    $valores_list.Add([PSCustomObject]@{
                        valor_nominal = $nom_val.ToString("F3", $culture)
                        maxima        = $max.ToString("F3", $culture)
                        minima        = $min.ToString("F3", $culture)
                        valor_atual   = $act_val.ToString("F3", $culture)
                    })
                }
                catch {
                    if (-not $modoRapido) {
                        Write-Host "[AVISO] Linha $i ignorada (erro de conversão): $($lines[$i])" #APENAS VERIFICAÇÃO
                    }
                    continue
                }
            }
            else {
                if (-not $modoRapido) {
                    Write-Host "[AVISO] Linha $i ignorada (formato inválido): $($lines[$i])" #APENAS VERIFICAÇÃO
                }
            }
        }

        #GERA ARQUIVO CSV DE SAÍDA
        $global:caminhoSaida = "C:\Users\oper\Desktop\DELETAR\dados.csv"

        
        #LINHA DE TESTE (NÃO ESTÁ SENDO USADA)
        #######################################################
        #$csvContent = $valores_list | ConvertTo-Csv -NoTypeInformation -UseCulture:$false
        #$csvContent | Out-File -FilePath $caminhoSaida -Encoding UTF8
        #############################################

        $valores_list | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $caminhoSaida

        return $null
    }
    catch {
        Write-Host "[ERRO] Falha no processamento: $_" #APENAS VERIFICAÇÃO
        return $null
    }
}


#########################################################################################################

# FUNÇÃO QUE LÊ UM ARQUIVO TXT E TRATA AS INFORMAÇÕES
function TratarArquivoCSV {
    param (
        [string]$caminhoArquivo,
        [bool]$modoRapido = $false
    )

    try {
        #LÊ TODAS AS LINHAS DO CSV COMO TEXTO BRUTO
        $linhas = Get-Content $caminhoArquivo

        if ($linhas.Count -lt 13) {
            #Write-Host "[AVISO] Arquivo possui menos de 13 linhas. Verifique o conteúdo." #APENAS VERIFICAÇÃO // TRATAMENTO DE ERRO
            return $null
        }

        #EXTRAI O NÚMERO DO ITEM DA CÉLULA A3 (LINHA 2 NO ÍNDICE, APÓS A VÍRGULA)
        $linhaA3 = $linhas[2]
        $global:numero_item = $linhaA3.Split(",")[1].Trim()

        # LÊ O CSV A PARTIR DA LINHA 12 (ÍNDICE 12 - A13 EM 1-BASED)
        $conteudoCsv = $linhas[12..($linhas.Count - 1)] -join "`n" | ConvertFrom-Csv

        # LISTA PARA OS DADOS PROCESSADOS
        $valores_list = [System.Collections.Generic.List[PSCustomObject]]::new()
        $culture = [System.Globalization.CultureInfo]::InvariantCulture

        foreach ($linha in $conteudoCsv) {
            try {
                $design_val  = [math]::Abs([double]$linha.'design val.')
                $upper_limit = [math]::Abs([double]$linha.'upper limit')
                $lower_limit = [math]::Abs([double]$linha.'lower limit')
                $mes_value   = [math]::Abs([double]$linha.'mes. value')

                $max = $design_val + $upper_limit
                $min = $design_val - $lower_limit

                $valores_list.Add([PSCustomObject]@{
                    valor_nominal = $design_val.ToString("F3", $culture)
                    maxima        = $max.ToString("F3", $culture)
                    minima        = $min.ToString("F3", $culture)
                    valor_atual   = $mes_value.ToString("F3", $culture)
                })
            }
            catch {
                if (-not $modoRapido) {
                    Write-Host "[AVISO] Linha ignorada por erro de conversão: $($_.Exception.Message)" #APENAS VERIFICAÇÃO
                }
            }
        }

        
        # Gera o arquivo CSV de saída
        $global:caminhoSaida = "C:\Users\oper\Desktop\DELETAR\dados.csv"

        # Sobrescreve o mesmo arquivo CSV com os dados processados
        $valores_list | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $caminhoSaida

        return $null
    }
    catch {
        Write-Host "[ERRO] Falha no processamento: $_" #APENAS VERIFICAÇÃO
        return $null
    }
}



################################################################################################################################################################

# FUNÇÃO PARA EXECUTAR A QUERY NA VIEW ITEMSTANDARDSBYTECHSTAGE_V E EXTRAIR AS COLUNAS DE VALOR NOMINAL, MÁXIMO, MÍNIMO E ID
function QuerySQL {
    param (
        [string]$item_code,
        [string[]]$tech_stage,
        [string]$global:outputFilePath = "C:\Users\oper\Desktop\DELETAR\output.csv"
    )

    # STRING DE CONEXÃO
    $connectionString = "Provider=sqloledb;Data Source=10.71.0.18;Initial Catalog=Manti_JBData;User ID=mantiapp;Password=nbyhtpp;"

    # VIEW A CONSULTAR
    $view = "dbo.ItemStandardsByTechStage_V"

    # ABRIR CONEXÃO
    $connection = New-Object -ComObject ADODB.Connection
    $connection.Open($connectionString)


    #MONTA A CLÁUSULA WHERE COM MÚLTIPLOS TECH_STAGE USANDO OR
    $whereStages = $tech_stage | ForEach-Object { "P9Stg = '$_'" }
    $whereClause = $whereStages -join " OR "

    $query = @"
        SELECT PropertyId, IsdOptVal, IsdMaxVal, IsdMinVal
        FROM $view
        WHERE ItemCode = '$item_code'
          AND ($whereClause)
"@

    $recordset = New-Object -ComObject ADODB.Recordset

    try {
        $recordset.Open($query, $connection)

        if ($recordset.EOF) {
            Write-Host "[AVISO] Nenhum registro encontrado" #APENAS VERIFICAÇÃO
            return $null
        } else {
            $resultados = @()

            while (!$recordset.EOF) {
                $propertyId = $recordset.Fields.Item("PropertyId").Value
                $isdOptVal = $recordset.Fields.Item("IsdOptVal").Value
                $isdMaxVal = $recordset.Fields.Item("IsdMaxVal").Value
                $isdMinVal = $recordset.Fields.Item("IsdMinVal").Value

                $resultados += [PSCustomObject]@{
                    PropertyId = $propertyId
                    IsdOptVal  = $isdOptVal
                    IsdMaxVal  = $isdMaxVal
                    IsdMinVal  = $isdMinVal
                }

                $recordset.MoveNext()
            }

            $resultados | Export-Csv -Path $outputFilePath -NoTypeInformation
            Write-Host "[INFO] Resultados exportados para $outputFilePath" #APENAS VERIFICAÇÃO
            return
        }

        $recordset.Close()
    } catch {
        Write-Host "[ERRO] Erro ao consultar $view : $_" #APENAS VERIFICAÇÃO
        return $null
    }

    #FECHAR CONEXÃO
    $connection.Close()
}




################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################

# FUNÇÃO PARA ENCONTRAR A WORK ORDER ID
function get_order {
    param (
        [string]$order,
        [string[]]$operation
    )

    #PARÂMETROS DE CONEXÃO
    $connectionString = "Provider=sqloledb;Data Source=10.71.0.18;Initial Catalog=Manti_JBData;User ID=mantiapp;Password=nbyhtpp;"
    
    #INICIALIZA A CONEXÃO
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)

    try {
        #ABRE A CONEXÃO
        $connection.Open()

        # CONSTRÓI A CLÁUSULA WHERE DINÂMICA COM MÚLTIPLAS OPERAÇÕES
        $likeClauses = $operation | ForEach-Object { "Work_Orders.WorkOrderOperDesc LIKE '%$_%'" }
        $whereClause = $likeClauses -join " OR "

        #MONTA A CONSULTA SQL
        $sql = @"
            SELECT TOP 1 Work_Orders.WorkOrderId
            FROM Manti_JBData.dbo.Production_Orders Production_Orders
            JOIN Manti_JBData.dbo.Work_Orders Work_Orders 
            ON Production_Orders.ProdOrderId = Work_Orders.ProdOrderId 
            WHERE Production_Orders.ProdOrderNo = '$order' 
              AND ($whereClause)
"@

        #CRIAR COMANDO SQL
        $command = $connection.CreateCommand()
        $command.CommandText = $sql

        #EXECUTA A CONSULTA E TENTA OBTER UM RESULTADO
        $reader = $command.ExecuteReader()
        if ($reader.Read()) {
            #ARMAZENA APENAS A WORK ORDER ID
            $global:result = $reader.GetValue(0)

            #FECHA O LEITOR E A CONEXÃO ANTES DE RETORNAR
            $reader.Close()
            $connection.Close()

            Write-Host "[INFO] WorkOrderId encontrado: $result" #APENAS VERIFICAÇÃO
            return $result
        } else {
            $reader.Close()
            $connection.Close()
            Write-Host "[INFO] Nenhum WorkOrderId encontrado." #APENAS VERIFICAÇÃO
            return $null
        }

    } catch {
        Write-Host "[ERRO] Erro ao consultar Work_Orders: $($_.Exception.Message)" -ForegroundColor Red #APENAS VERIFICAÇÃO
        return $null
    } finally {
        if ($connection.State -eq "Open") {
            $connection.Close()
        }
    }
}


################################################################################################################################################################

#FUNÇÃO PARA MESCLAR OS DADOS E CRIAR A TABELA PARA INSERÇÃO NO MANTI
function mesclar_dados {
    param (
        [string]$caminho_dados,
        [string]$caminho_output 
    )

    #LÊ OS ARQUIVOS CSV
    $dados = Import-Csv -Path $caminho_dados
    $output = Import-Csv -Path $caminho_output

    #LISTA PARA ARMAZENAR OS RESULTADOS
    $resultados = @()

    foreach ($linha_dado in $dados) {
        #ARREDONDA OS VALORES DO ARQUIVO DADOS.CSV PARA 2 CASAS DECIMAIS
        $val_nom = [math]::Round([double]$linha_dado.valor_nominal, 2)
        $val_max = [math]::Round([double]$linha_dado.maxima, 2)
        $val_min = [math]::Round([double]$linha_dado.minima, 2)

        foreach ($linha_output in $output) {
            #ARREDONDA OS VALORES DO ARQUIVO OUTPUT.CSV PARA 2 CASAS DECIMAIS
            $out_nom = [math]::Round([double]$linha_output.IsdOptVal, 2)
            $out_max = [math]::Round([double]$linha_output.IsdMaxVal, 2)
            $out_min = [math]::Round([double]$linha_output.IsdMinVal, 2)

            #FAZ A COMPARAÇÃO COM BASE NOS VALORES ARREDONDADOS
            if (
                ($val_nom -eq $out_nom) -and
                ($val_max -eq $out_max) -and
                ($val_min -eq $out_min)
            ) {
                $resultados += [PSCustomObject]@{
                    PropertyId    = $linha_output.PropertyId
                    valor_atual = $linha_dado.valor_atual
                }
                break #SE JÁ ENCONTROU CORRESPONDÊNCIA NÃO PRECISA MAIS VERIFICAR
            }
        }
    }

    #DEFINE O CAMINHO DE SAÍDA
    $global:caminho_manti = "C:\Users\oper\Desktop\DELETAR\manti_values.csv"

    #EXPORTA OS RESULTADOS
    $resultados | Export-Csv -Path $caminho_manti -NoTypeInformation -Encoding UTF8

    Write-Host "[INFO] Arquivo mesclado criado com sucesso em: $caminho_manti"
}


################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################


function Send_to_Manti {
    param (
        [string]$workOrderId,        #ID DE ORDEM DE TRABALHO OBTIDO PELA FUNÇÃO GET_ORDER
        [string]$mantiValuesPath,    #CAMINHO DO ARQUIVO MANTI_VALUES.CSV
        [int]$n_peca                 #NÚMERO DA PEÇA (SAMPLETESTRESULT ESPECIAL PARA PROPERTY ID 5730)
    )

    #CONEXÃO COM O BANCO DE DADOS
    $connectionString = "Provider=sqloledb;Data Source=10.71.0.18;Initial Catalog=Manti_JBData;User ID=mantiapp;Password=nbyhtpp;"
    $connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $connection.Open()

    try {
        #INSERIR O NÚMERO DA PEÇA COM O ID FIXO = 5730
        $ID_peca = 5730
        $sql_peca = @"
        INSERT INTO Manti_JBData.dbo.Machine_Gelem_Result
        (WorkOrderId, SampleItemNo, PropertyId, SampleTestResult, DateLastModified, UserLastModified, DateCurrentUsed, MachineType, SinturRunNo) 
        VALUES ('$workOrderId', NULL, '$ID_peca', '$n_peca', GETDATE(), 0, NULL, 5, NULL) 
"@
        $command = $connection.CreateCommand()
        $command.CommandText = $sql_peca
        $command.ExecuteNonQuery() | Out-Null

        #LÊ O CSV GERADO PELA FUNÇÃO MESCLAR_DADOS
        $dadosCsv = Import-Csv -Path $mantiValuesPath

        foreach ($linha in $dadosCsv) {
            $PropertyId = $linha.PropertyId
            $SampleTestResult = $linha.valor_atual  

            $sql = @"
            INSERT INTO Manti_JBData.dbo.Machine_Gelem_Result 
            (WorkOrderId, SampleItemNo, PropertyId, SampleTestResult, DateLastModified, UserLastModified, DateCurrentUsed, MachineType, SinturRunNo) 
            VALUES ('$workOrderId', NULL, '$PropertyId', '$SampleTestResult', GETDATE(), 0, NULL, 5, NULL)
"@
            $command = $connection.CreateCommand()
            $command.CommandText = $sql
            $command.ExecuteNonQuery() | Out-Null
        }

        Write-Host "[INFO] Todos os dados do arquivo foram inseridos com sucesso."

    } catch {
        Write-Host "[ERRO] Erro ao inserir no banco de dados: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        $connection.Close()
    }
}



################################################################################################################################################################


#CONFIGURAÇÃO PRINCIPAL - DEFINE O CAMINHO DA PASTA A SER MONITORADA E O DESTINO DA PASTA PARA QUAL SERÁ ENVIADO OS ARQUIVOS APÓS PROCESSADOS.
$monitorarPasta = "G:\MONITORAR"
$destinoPasta = "C:\Users\oper\Desktop\DELETAR"


# LIMPEZA DE EVENTOS E SUBSCRIBERS ANTERIORES PARA EVITAR DUPLICIDADE AO REEXECUTAR O SCRIPT
try {
    Unregister-Event -SourceIdentifier * -ErrorAction SilentlyContinue
    Get-EventSubscriber | Unregister-Event -Force -ErrorAction SilentlyContinue
} catch {
    Write-Host "[INFO] Nenhum evento anterior para remover."
}



################################################################################################################################################


# WATCHER 
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $monitorarPasta
$watcher.Filter = "*.*"  # MONITORAR TODOS OS ARQUIVOS
$watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true


################################################################################################################################################


# AÇÃO QUANDO O ARQUIVO É CRIADO
$action = {
    param($source, $eventArgs)

    #DESABILITA EVENTOS ENQUANTO PROCESSA O ARQUIVO
    $watcher = $source
    $watcher.EnableRaisingEvents = $false

    try {
        $arquivoCriado = $eventArgs.FullPath
        $extensao = [System.IO.Path]::GetExtension($arquivoCriado).ToLower()

        # VERIFICAR SE O ARQUIVO É .TXT OU .CSV
        if ($extensao -ne ".txt" -and $extensao -ne ".csv") {
            Write-Host "Arquivo ignorado: $arquivoCriado" #TRATAMENTO DE ERRO
            return
        }

        $nomeArquivo = [System.IO.Path]::GetFileNameWithoutExtension($arquivoCriado)
        $destinoPasta = $event.MessageData.destinoPasta

        if (-not (AguardarArquivo -caminhoArquivo $arquivoCriado)) {
            Write-Host "Erro: Nao foi possivel acessar o arquivo '$arquivoCriado'." #TRATAMENTO DE ERRO
            return
        }

        try {
            $n_wo_match = $nomeArquivo -match "^(\d{7})"
            $n_wo = $matches[1]

            $nomePadronizado = $nomeArquivo -replace '[-_]', '-'
            $partes = $nomePadronizado -split '-'

            if ($partes.Count -lt 3) {
                Write-Host "Erro: Nome do arquivo '$nomeArquivo' nao esta no formato esperado." #TRATAMENTO DE ERRO
                return
            }

            $n_peca = $partes[1] -replace '[^0-9]', ''
            $maquina = ($nomePadronizado -split '-', 3)[2].ToUpper()

            Write-Host "WO: $n_wo, PC: $n_peca, MAQUINA: $maquina" #APENAS VERIFICAÇÃO

        } catch {
            Write-Host "Erro ao processar o nome do arquivo: $_" #TRATAMENTO DE ERRO
            return
        }

        $destino = MoverArquivo -sourcePath $arquivoCriado -destinoPasta $destinoPasta
        if (-not $destino) {
            Write-Host "Erro ao mover arquivo" #TRATAMENTO DE ERRO
            return
        }

        if (-not (AguardarArquivo -caminhoArquivo $destino)) {
            Write-Host "Erro: Arquivo movido nao esta acessivel." #TRATAMENTO DE ERRO
            return
        }

        
        # CHAMAR FUNÇÕES
        
        if ($extensao -eq '.txt') {
            TratarArquivoTXT -caminhoArquivo $destino
        }
        elseif ($extensao -eq '.csv') {
            TratarArquivoCSV -caminhoArquivo $destino
        }
        else {
            Write-Host "Arquivo ignorado: $arquivoCriado" #TRATAMENTO DE ERRO
        }


################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################
# TRECHO DO CÓDIGO PARA DEFINIR O CENTRO DE CUSTO ATRAVÉS DO NOME NO RELATÓRIO

        #NORMALIZA A STRING PARA FACILITAR A COMPARAÇÃO
        $maquina_normalizada = $maquina.Trim().ToLower()
        Write-Host $maquina_normalizada #APENAS VERIFICAÇÃO

        #POSSÍVEIS MÁQUINAS
        $grupoANCA_MX7 = @( "anca mx7", "ancamx7", "anca_mx7", "anca-mx7", "mx7", "wg3", "wg4" )   # B11
        $grupoANCA_CPX = @( "ancacpx", "anca-cpx", "anca_cpx", "cpx", "anca cpx", "cp1" )          # 218, 219
        $grupoHAAS_1   = @( "haas 1", "haas1", "haas_1", "haas-1", "hg1" )                         # A08, A12
        $grupoHAAS_2   = @( "haas 2", "haas2", "haas_2", "haas-2", "hg2" )                         # A08
        $grupoHAAS_3   = @( "haas 3", "haas3", "haas_3", "haas-3", "hg3" )                         # B11, A08
        $grupoHAAS_4   = @( "haas 4", "haas4", "haas_4", "haas-4", "haas cu", "haascu", "haas_cu", "haas-cu", "cu", "cu1" ) # A08
        $grupoWALTER   = @( "walter", "walter 1", "walter 2", "walter-1", "walter-2", "walter_1", "walter_2", "walter1", "walter2", "wg1", "wg2" ) # B11
        $grupoCONTROLE = @( "controle" )                                                            # 901

        #VERIFICA A QUAL GRUPO A MÁQUINA PERTENCE E ATRIBUI OS TECH STAGES
        if ($grupoANCA_MX7 -contains $maquina_normalizada) {
            $stage = @("B11")
            $operacao = @( "SPIRAL GRINDING", "GRINDING WALTER")
        }
        elseif ($grupoANCA_CPX -contains $maquina_normalizada) {
            $stage = @("218", "219")
            $operacao = @( "O.D GRID 10?BT", "O.D GRID. 90NBT" )
        }
        elseif ($grupoHAAS_1 -contains $maquina_normalizada) {
            $stage = @("A08", "A12")
            $operacao = @( "INSERT GRIND 5E", "CHIP SURFACE GR", "INSERT GRIND 6E")
        }
        elseif ($grupoHAAS_2 -contains $maquina_normalizada) {
            $stage = @("A08")
            $operacao = @( "INSERT GRIND 6E")
        }
        elseif ($grupoHAAS_3 -contains $maquina_normalizada) {
            $stage = @("B11", "A08")
            $operacao = @("GRINDING WALTER", "INSERT GRIND 5E")
        }
        elseif ($grupoHAAS_4 -contains $maquina_normalizada) {
            $stage = @("A08")
            $operacao = @( "INSERT GRIND CU")
        }
        elseif ($grupoWALTER -contains $maquina_normalizada) {
            $stage = @("B11")
            $operacao = @( "GRINDING WALTER")
        }
        elseif ($grupoCONTROLE -contains $maquina_normalizada) {
            $stage = @("901")
            $operacao = @( "CHECK DIMENSION")
        }
        else {
            Write-Host "[AVISO] Máquina '$maquina' não reconhecida."
            $stage = @("DESCONHECIDO")
            $operacao = @( "DESCONHECIDO")
        }
################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################################


    #CHAMA A FUNÇÃO QUERY SQL    
    QuerySQL -item_code $numero_item -tech_stage $stage

    #CHAMA A FUNÇÃO PARA EXTRAIR A WORK ORDER ID
    get_order -order $n_wo -operation $operacao

    #CHAMA A FUNÇÃO PARA MESCLAR OS DADOS E CRIAR A TABELA PARA INSERÇÃO NO MANTI
    mesclar_dados -caminho_dados $caminhoSaida -caminho_output $outputFilePath

    #CHAMA A FUNÇÃO PARA INSERIR OS DADOS NO MANTI
    Send_to_Manti -workOrderId $result -mantiValuesPath $caminho_manti -n_peca $n_peca

    

    }
    finally {
        #REABILITA EVENTOS
        $watcher.EnableRaisingEvents = $true
    }
}


##############################################################################################################################################

#REGISTRAR O EVENTO PASSANDO O CAMINHO DE DESTINO COMO MessageData
Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $action -MessageData @{ destinoPasta = $destinoPasta }

#############################################################################################################################################

# MANTÉM O SCRIPT RODANDO
Write-Host "Monitorando a pasta: $monitorarPasta" #APENAS VERIFICAÇÃO
while ($true) {
    Start-Sleep -Seconds 2
}

################################################################################################################################################################











