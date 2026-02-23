function app 
{
    param(
        [String]$UserAdmin = "master.user"
    )

    cd "C:\Users\${UserAdmin}"
    Import-Module ActiveDirectory
    
    $formApp = formGen(@("RPA EMPRESA",
    400,
    400))

    $labelApp = labelGen -msg "Seja bem-vindo de volta, é um prazer tê-lo aqui."

    $formApp.Controls.Add($labelApp)

    $btnOpUnitaria = createBtnAct -acao ${function:operationOneUser} -nameBtn "Operações Unitárias" -dimensoes @(20, 80) -paramAct ""
    $formApp.Controls.Add($btnOpUnitaria)

    $btnOpPlural = createBtnAct -acao ${function:operationMoreUsers} -nameBtn "Operações Plurais" -dimensoes @(20, 130) -paramAct ""
    $formApp.Controls.Add($btnOpPlural)

    $btnOpKace = createBtnAct -acao ${function:operationFormKace} -nameBtn "Operações Kace" -dimensoes @(20, 180) -paramAct ""
    $formApp.Controls.Add($btnOpKace)

    $btnOpExit = createBtnAct -acao ${function:encerrar} -nameBtn "Encerrar" -dimensoes @(20, 230) -paramAct ""
    $formApp.Controls.Add($btnOpExit)

    $formApp.Topmost = $true

    $formApp.ShowDialog()
}

function operationOneUser
{
    $userAd = formInput -message "Informe o(a) usuário(a) (Ex.: claudio.bernardo):"

    if($userAd -eq $null)
    {
        $entry = @("Nenhum usuário informado.", "Exceção", 'Error')
        alertMessage($entry)
        return
    }

    if (isUserAd($userAd))
    {
        $entry = @("O usuário ${userAd} existe.", "Sucesso", 'Information')
        alertMessage($entry)
    }
      else
    {
        $entry = @("O usuário ${userAd} não existe ou está incorreto.", "Exceção", 'Error')
        alertMessage($entry)
        operationOneuser("")
    }

    Get-ADUser -Identity $userAd
    
    $userOrgUnit = getUserOrgUnit($userAd)
    $userGroups = getUserGroups($userAd)

    $acoes = @(
        "Reset de Senha"
        "Adicionar à impressora"
        "Remover à impressora"
        "Informar outro usuário"
        "Sair"
    )

    $action = Show-OptionForm -optns $acoes

    switch ($action) {
        "Reset de Senha" {
            return actionResetPassword("")
        }

        "Adicionar à impressora" {
            return actionAddUserGroup($userAd)
        }

        "Remover à impressora" {
            return actionRemoveUserGroup($userAd)
        }

        "Informar outro usuário" {
            return operationOneUser("")
        }
        "Sair" {
            return 
        }
       }

    return 
}

function operationMoreUsers()
{
    Import-Csv C:\Users\master.user\Documents\ADUser.csv | ForEach-Object {
        $userAd = $_."samAccountName"
        $dataUserAd = Get-ADUser -Identity $userAd
    }
    $opcoes = @(
        "Reset de Senha Padrão"
        "Reset de Senha S/ Conexão"
        "Verificar lista de usuários"
        "Verificar grupo de usuários"
        "Remover usuários de um Grupo"
        "Adicionar usuários à um grupo"
        "Atualizar Termos Kace"
        "Sair"
    )

    $optn = Show-OptionForm -titulo "Qual opção deseja realizar?" -optns $opcoes

    if (-not $optn) {
        return
    }

    switch ($optn) {

        "Reset de Senha Padrão" {
            Import-Csv C:\Users\master.user\Documents\ADUser.csv | ForEach-Object {
                $userAd = $_."samAccountName"
                alterarSenha($userAd)
            }
            operationMoreUsers("")
        }

        "Reset de Senha S/ Conexão" {
            
            Import-Csv C:\Users\master.user\Documents\ADUser.csv | ForEach-Object {
                $userAd = $_."samAccountName"
                $password = (formPassword -text "Digite sua nova senha ${userAd}: ") | Where-Object { $_ -is [string] -and -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 1
                alterarSenhaOffline($userAd, $password)
            }
            operationMoreUsers("")
        }

        "Verificar lista de usuários" {
            Start-Process "notepad.exe" "C:\Users\master.user\Documents\ADUser.csv"
            operationMoreUsers("")
        }

        "Verificar grupo de usuários" {
            Import-Csv C:\Users\master.user\Documents\ADUser.csv | ForEach-Object {
                $userAd = $_."samAccountName"
                getUserGroups($userAd)
            }
            operationMoreUsers("")
        }

        "Remover usuários de um Grupo" {
            Import-Csv C:\Users\master.user\Documents\ADUser.csv | ForEach-Object {
                $userAd = $_."samAccountName"
                actionRemoveUserGroup($userAd)
            }
            operationMoreUsers("")
        }

        "Adicionar usuários à um grupo" {
            Import-Csv C:\Users\master.user\Documents\ADUser.csv | ForEach-Object {
                $userAd = $_."samAccountName"
                actionAddUserGroup($userAd)
            }
        }

        "Atualizar Termos Kace" {
                $data_form = iteradorTermos -file "C:\Users\master.user\Downloads\termos_atualizar.xlsx"
                $totalLinhas = $data_form["USUARIO"].Count

                for ($i = 0; $i -lt $totalLinhas; $i++) {
                    
                                       
                    if ($data_form["MODELO"][$i] -match '\d{4}(?!.*\d)') {
                        $modelo = $matches[0]
                    } else {
                        $modelo = ""
                    }
                    
                    $perifericos = @()

                    if ($data_form["PERIFERICOS"][$i]) {
                        $perifericos = $data_form["PERIFERICOS"][$i] -split '\s*,\s*'
                    }
                    
                    $patrimonio = ""
                    $patriMon1 = ""
                    $patriMon2 = ""

                    if ($data_form["PATRIMONIO"][$i])
                    {
                        $patrimonio = '(' + $data_form["DISPOSITIVO"][$i] + ')' + $data_form["PATRIMONIO"][$i]
                    }

                    if ($data_form["PATRIMONIO M1"][$i])
                    {
                        $patriMon1 = '(' + $data_form["MONITOR 1 - TAG"][$i] + ')' + $data_form["PATRIMONIO M1"][$i]
                    } 

                    if ($data_form["PATRIMONIO M2"][$i])
                    {
                        $patriMon2 = '(' + $data_form["MONITOR 2 - TAG"][$i] + ')' + $data_form["PATRIMONIO M2"][$i]
                    }

                    $modeloM1, $tagM1 = Get-MonitorInfo $data_form["MONITOR 1 - TAG"][$i]
                    $modeloM2, $tagM2 = Get-MonitorInfo $data_form["MONITOR 2 - TAG"][$i]

                    # Arrays (ou strings) para o termo
                    $arrModeloMonitor = @($modeloM1, $modeloM2) | Where-Object { $_ -ne "" }
                    $arrServiceTagMonitor = @($tagM1, $tagM2) | Where-Object { $_ -ne "" }
                    $arrMonPatri = @($($data_form["PATRIMONIO M1"][$i]), $($data_form["PATRIMONIO M2"][$i]))

                    $UserAdmin = $env:USERNAME

                    
                    scriptTermo @(
                        $data_form["DISPOSITIVO"][$i],
                        $data_form["MODELO"][$i], 
                        $data_form["SERVICE_TAG"][$i],
                        $data_form["USUARIO"][$i],
                        $data_form["STATUS"][$i],
                        $data_form["LOCALIDADE"][$i],
                        $data_form["SETOR"][$i],
                        $data_form["MONITOR 1 - TAG"][$i],
                        $data_form["MONITOR 2 - TAG"][$i],
                        $data_form["DOCKSTATION - TAG"][$i],
                        $patrimonio,
                        $patriMon1,
                        $patriMon2,
                        "(Adicionar o chamado par aqui)",
                        "Atualização de Termos KACE com base na AUDITORIA"
                    )


                    $data_termo = @(
                        $data_form["DISPOSITIVO"][$i],
                        $modelo,
                        $data_form["SERVICE_TAG"][$i],
                        $data_form["USUARIO"][$i],
                        $data_form["STATUS"][$i],
                        $data_form["LOCALIDADE"][$i],
                        $UserAdmin.Replace(".ope", ""),
                        ($arrModeloMonitor -join ','),  
                        ($arrServiceTagMonitor -join ','),  
                        ($arrMonPatri -join ','),  
                        $data_form["PATRIMONIO"][$i],
                        $data_form["DOCKSTATION - TAG"][$i],
                        ($perifericos -join ',')
                    )

                    criadorTermoKace -dataTermo $data_termo

                }
            }
        
        "Sair" {
            return 
        }

       }

    return 
}

function operationFormKace()
{
    $opcoesFormKace = @(
        "Criação KACE"
        "Alteração KACE"
        "Reparo Dell"
        "Sair"
    )

    $optnFormKace = Show-OptionForm -titulo "Qual opção KACE deseja?" -optns $opcoesFormKace

    switch ($optnFormKace) {
        "Criação KACE"     { actionCriacaoKace("") }
        "Alteração KACE"   { actionAlteracaoKace("") }
        "Reparo Dell"      { actionReparoDell("") }
        "Sair" { return }
    }

    return
}

function actionAlteracaoKace()
{
    $opcoes = @(
        "DESKTOP"
        "NOTEBOOK"
        "MONITOR"
        "DOCKSTATION"
        "Voltar para o início"
        "Sair"
    )

    $tipoAlteracao = Show-OptionForm -titulo "Informe o dispositivo:" -optns $opcoes

    switch ($tipoAlteracao) {
        "DESKTOP" {
            $dados = modeloDeskAtivo("")
            Write-Host "Tipo: " $dados[1] "Modelo: " $dados[2] "Servicetag: " $dados[3]
            formAlteracaoKace -tipo $dados[1] -model $dados[2] -serviceTag $dados[3]
                    }
        "NOTEBOOK" {
            $dados = modeloNoteAtivo("")
            Write-Host "Tipo: " $dados[1] "Modelo: " $dados[2] "Servicetag: " $dados[3]
            formAlteracaoKace -tipo $dados[1] -model $dados[2] -serviceTag $dados[3]
        }
        "MONITOR" {
            $dados = modeloMonAtivo("")
            Write-Host "Tipo: " $dados[0] "Modelo: " $dados[1] "Servicetag: " $dados[2]
            formAlteracaoKace -tipo $dados[0] -model $dados[1] -serviceTag $dados[2]
        }
        "DOCKSTATION" {
            $dados = modeloDockAtivo("")
            formAlteracaoKace -tipo $dados[0] -model $dados[1] -serviceTag $dados[2] 
        }
        "Voltar para o início" {
            operationFormKace("")
        }
        "Sair" {
            return 
        }
       }

    return 
}

function actionCriacaoKace()
{
    $opcoes = @(
        "DESKTOP"
        "NOTEBOOK"
        "MONITOR"
        "DOCKSTATION"
        "Voltar para o início"
        "Sair"
    )

    $tipoAlteracao = Show-OptionForm -titulo "Informe o dispositivo:" -optns $opcoes

    switch ($tipoAlteracao) {
        "DESKTOP" {
            $dados = modeloDeskAtivo("")
            
            formCriacaoKace -tipo $dados[1] -model $dados[2] -serviceTag $dados[3]
        }
        "NOTEBOOK" {
            $dados = modeloNoteAtivo("")
            formCriacaoKace -tipo $dados[1] -model $dados[2] -serviceTag $dados[3]
        }
        "MONITOR" {
            $dados = modeloMonAtivo("")
            formCriacaoKace -tipo $dados[0] -model $dados[1] -serviceTag $dados[2]
        }
        "DOCKSTATION" {
            $dados = modeloDockAtivo("")
            formCriacaoKace -tipo $dados[0] -model $dados[1] -serviceTag $dados[2]
        }
        "Voltar para o início" {
            operationFormKace("")
        }
        "Sair" {
            return 
        }
       }

    return 
}

function actionReparoDell()
{
    $opcoes = @(
        "DESKTOP"
        "NOTEBOOK"
        "MONITOR"
        "DOCKSTATION"
        "Voltar para o início"
        "Sair"
    )

    $tipoAlteracao = Show-OptionForm -titulo "Informe o dispositivo:" -optns $opcoes

    switch ($tipoAlteracao) {
        "DESKTOP" {
            $dados = modeloDeskAtivo("")
            formReparoDell -tipo $dados[0] -model $dados[1] -serviceTag $dados[2]
        }
        "NOTEBOOK" {
            $dados = modeloNoteAtivo("")
            formReparoDell -tipo $dados[0] -model $dados[1] -serviceTag $dados[2]
        }
        "MONITOR" {
            $dados = modeloMonAtivo("")
            formReparoDell -tipo $dados[0] -model $dados[1] -serviceTag $dados[2]
        }
        "DOCKSTATION" {
            $dados = modeloDockAtivo("")
            formReparoDell -tipo $dados[0] -model $dados[1] -serviceTag $dados[2]
        }
        "Voltar para o início" {
            operationFormKace("")
        }
        "Sair" {
            return 
        }
       }

    return 
}

function actionResetPassword()
{
    $opcoes = @(
        "Conectado à Internet"
        "Sem conexão à internet"
        "Sair"
    )

    $userAdmin = $env:USERNAME
    $userAdmin = $userAdmin.Replace(".ope", "")

    New-Item -Path "C:\Users\master.user\Documentos\LOG_ABERTURA_CHAMADOS\ABERTURA_CHAMADO_${userAd}_RESET_SENHA.txt" -ItemType "file" -Value "
LOGIN: ${userAd}
SERVICE TAG:
DESC: Por favor, preciso que minha senha seja resetada.
A/C: ${userAdmin}" -Force

    $acessoNet = Show-OptionForm -titulo "O usuário está..." -optns $opcoes
    switch($acessoNet) 
    {
        "Conectado à Internet" { 
                alterarSenha -user $userAd
            }
        "Sem conexão à internet" { 
                $password = (formPassword -text "Digite sua senha:") | Where-Object { $_ -is [string] -and -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 1
                alterarSenhaOffline -user $userAd -pssw $password
            }
        "Sair" {
            return 
        }
      }
      
      Start-Process "notepad.exe" "C:\Users\master.user\Documentos\LOG_ABERTURA_CHAMADOS\ABERTURA_CHAMADO_${userAd}_RESET_SENHA.txt"

    return 
}

function actionAddUserGroup($userAd) 
{
    $entryAlert = @("Você estará adicionando o ${userAd} da determinada impressora que selecionar.", "Informativo", 'Information')
    alertMessage($entryAlert)
    $cidades = @(
        "CDD1"
        "CDD2"
        "CDD3"
        "CDD4"
        "Sair"
    )

    $optnCidade = Show-OptionForm -titulo "Selecione a cidade:" -optns $cidades
    switch ($optnCidade)
    {
        "CDD1" {
                    $setores = @(
                        "LOJA CDD1"
                        "RH CDD1"
                        "ALMOX CDD1"
                        "ADM CDD1"
                        "ADM CDD1 COLOR"
                        "LAB CDD1"
                        "Sair"
                    )
                    $setor = Show-OptionForm -titulo "Selecione o setor:" -optns $setores
                    switch($setor) 
                    {
                        "LOJA CDD1" { addGroup -group "IMP_EMP_CDD1_LOJA_PB" -user $userAd }
                        "RH CDD1" { addGroup -group "IMP_EMP_CDD1RH_PB" -user $userAd }
                        "ALMOX CDD1" { addGroup -group "IMP_EMP_CDD1_ALMOXARIFADO_PB" -user $userAd }
                        "ADM CDD1" { addGroup -group "IMP_EMP_CDD1_ADMINISTRATIVO_PB" -user $userAd }
                        "ADM CDD1 COLOR" { addGroup -group "IMP_EMP_CDD1_SFD_ADMINISTRATIVO_COLOR" -user $userAd }
                        "LAB CDD1" { addGroup -group "IMP_EMP_CDD1_ETA_LABORATORIO_PB" -user $userAd }
                        "Sair" { app("") }
                    }
            }
        "CDD2" {
                    $setores = @(
                        "LOJA CDD2"
                        "GSO CDD2"
                        "Sair"
                    )
                    $setor = Show-OptionForm -titulo "Selecione o setor:" -optns $setores
                    switch($setor)
                    {
                        "LOJA CDD2" { addGroup -group "IMP_EMP_CDD2_LOJA_PB" -user $userAd }
                        "GSO CDD2" { addGroup -group "IMP_EMP_CDD2_GSO_PB" -user $userAd }
                        "Sair" { app("") }
                    }
            }
        "CDD3" {
                $setores = @(
                        "LOJA CDD3"
                        "Sair"
                    )
                $setor = Show-OptionForm -titulo "Selecione o setor:" -optns $setores
                switch($setor)
                    {
                        "LOJA CDD3" { addGroup -group "IMP_EMP_CDD3_LOJA_PB" -user $userAd }
                        "Sair" { app("") }
                    }
            }
        "CDD4" {
                $setores = @(
                        "LOJA CDD4"
                        "Sair"
                    )
                $setor = Show-OptionForm -titulo "Selecione o setor:" -optns $setores
                switch($setor)
                    {
                        "LOJA CDD4" { addGroup -group "IMP_EMP_CDD4_LOJA_PB" -user $userAd }
                        "Sair" { app("") }
                    }
            }

        "Sair" {
            return 
        }
       }

    return 
}

function actionRemoveUserGroup($userAd)
{
    $entryAlert = @("Você estará removendo o ${userAd} da determinada impressora que selecionar.", "Informativo", 'Information')
    alertMessage($entryAlert)

    $cidades = @(
        "CDD1"
        "CDD2"
        "CDD3"
        "CDD4"
        "Sair"
    )

    $optnCidade = $setor = Show-OptionForm -titulo "Selecione a cidade:" -optns $cidades
    
    switch ($optnCidade)
    {
        "CDD1" {
                    $setores = @(
                        "LOJA CDD1"
                        "RH CDD1"
                        "ALMOX CDD1"
                        "ADM CDD1"
                        "ADM CDD1 COLOR"
                        "LAB CDD1"
                        "Sair"
                    )
                    $setor = Show-OptionForm -titulo "Selecione o setor:" -optns $setores
                    switch($setor) 
                    {
                        "LOJA CDD1" { rmGroup -group "IMP_EMP_CDD1_LOJA_PB" -user $userAd }
                        "RH CDD1" { rmGroup -group "IMP_EMP_CDD1_RH_PB" -user $userAd }
                        "ALMOX CDD1" { rmGroup -group "IMP_EMP_CDD1_ALMOXARIFADO_PB" -user $userAd }
                        "ADM SFD CDD1 COLOR" { rmGroup -group "IMP_EMP_CDD1_ADMINISTRATIVO_COLOR" -user $userAd }
                        "ADM SFD CDD1" { rmGroup -group "IMP_EMP_CDD1_ADMINISTRATIVO_PB" -user $userAd }
                        "Sair" { app("") }
                    }
            }
        "CDD2" {
                    $setores = @(
                        "LOJA CDD2"
                        "GSO CDD2"
                        "Sair"
                    )
                    $setor = Show-OptionForm -titulo "Selecione o setor:" -optns $setores
                    switch($setor)
                    {
                        "LOJA NAT" { rmGroup -group "IMP_EMP_CDD2_LOJA_PB" -user $userAd }
                        "GSO NAT" { rmGroup -group "IMP_EMP_CDD2_GSO_PB" -user $userAd }
                        "Sair" { app("") }
                    }
            }
        "CDD3" {
                    $setores = @(
                        "LOJA CDD3"
                        "Sair"
                    )

                    $setor = Show-OptionForm -titulo "Selecione o setor:" -optns $setores
                    switch($setor)
                    {
                        "LOJA CDD3" { rmGroup -group "IMP_EMP_CDD3_LOJA_PB" -user $userAd }
                        "Sair" { app("") }
                    }
                 }
        "CDD4" {
                    $setores = @(
                        "LOJA CDD4"
                        "Sair"
                    )

                    $setor = Show-OptionForm -titulo "Selecione o setor:" -optns $setores
                    switch($setor)
                    {
                        "LOJA CDD4" { rmGroup -group "IMP_EMP_CDD4_LOJA_PB" -user $userAd }
                        "Sair" { app("") }
                    }
                 }
        "Sair" {
                return
            }
     }
     return
}

function isDispositivoAd($entry)
{
    $dispositivo, $serviceTag = $entry

    switch($dispositivo) 
    {
        "DESKTOP"  { 
                    $result = Get-ADComputer -SearchBase "OU=NOTEBOOKS,OU=EMPRESA,OU=BRASIL,DC=DOMINIO,DC=LOCAL" -Filter { CN -eq $serviceTag }-ErrorAction SilentlyContinue
                    return $null -ne $result
                   }
        "NOTEBOOK" {
                    $result = Get-ADObject -SearchBase "OU=NOTEBOOKS,OU=EMPRESA,OU=BRASIL,DC=DOMINIO,DC=LOCAL" -Filter { CN -eq $serviceTag }-ErrorAction SilentlyContinue
                    return $null -ne $result
                   } 
    }
}

function isUserAd($user)
{
    $userAd = Get-ADUser -Filter { SamAccountName -eq $user } -ErrorAction SilentlyContinue
    return $null -ne $userAd
}

function iteradorTermos 
{
    param(
        [String]$file
    )

    $data = Import-Excel -Path $file -WorksheetName 'Planilha1' -StartRow 1

    if (-not $data) {
        Write-Host "Arquivo sem dados."
        return
    }

    $colunas = $data[0].PSObject.Properties.Name
    $data_form = @{}

    
    Write-Host "Número de linhas: $($data.Count)"
    Write-Host "Colunas: $($data[0].PSObject.Properties.Name -join ', ')"

    Write-Host "Primeira linha:"
    $data[0].PSObject.Properties | ForEach-Object {
        Write-Host "  $($_.Name) = $($_.Value)"
    }

    foreach ($coluna in $colunas) {
        Write-Host "`n--- Coluna: $coluna ---"
        $valoresColuna = @()

        foreach ($linha in $data) {
            $valor = $linha."$coluna"
            $valoresColuna += $valor
        }

        Write-Host "Valores coletados na coluna '$coluna':"
        if ($valoresColuna.Count -eq 0) {
            Write-Host "(nenhum valor encontrado)"
        }
        else {
            for ($i = 0; $i -lt $valoresColuna.Count; $i++) {
                Write-Host "  Linha $($i + 1): '$($valoresColuna[$i])'"
            }
        }

        $data_form[$coluna] = $valoresColuna
    }
    return $data_form
}

function scriptTermo 
{
    param(
    [String[]] $argumentos
    )

    $argumentos | ForEach-Object { Write-Host $_ }

    $tipo = $argumentos[0]

    if ($tipo -eq "DESKTOP" -or $tipo -eq "NOTEBOOK")
    {
        
        $model = $argumentos[1]
        $serviceTag = $argumentos[2]
        $usuario = $argumentos[3]
        $status = $argumentos[4]
        $localidade = $argumentos[5]
        $setor = $argumentos[6]
        $monitor1Tag = $argumentos[7]
        $monitor2Tag = $argumentos[8]
        $dockstationTag = $argumentos[9]
        $patrimonio = $argumentos[10]
        $patriMonitor1 = $argumentos[11]
        $patriMonitor2 = $argumentos[12]
        $chamadoPar = $argumentos[13]
        $observacao = $argumentos[14]
        
        $filePath = "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\ALTERACAO_KACE_${usuario}_${model}_${serviceTag}_${status}.txt"
        
        $conteudo = @"
TIPO: ${tipo}
SERVICE TAG: ${model}_${serviceTag}
USUARIO: ${usuario}
STATUS DO ATIVO: ${status}
LOCAL DO ATIVO: ${localidade}
DEPARTAMENTO: ${setor}
MONITOR: ${monitor1Tag} ${monitor2Tag}
DOCKSTATION: ${dockstationTag}
PATRIMONIO: ${patrimonio} ${patriMonitor1} ${patriMonitor2}
CHAMADO PAR: ${chamadoPar}
OBSERVAÇÃO: ${observacao}
"@

        New-Item -Path $filePath -ItemType "File" -Value $conteudo -Force
        Start-Process "notepad.exe" $filePath 
       }
    if ($tipo -eq "MONITOR")
    {
        $model = $argumentos[1]
        $serviceTag = $argumentos[2]
        $usuario = $argumentos[3]
        $status = $argumentos[4]
        $localidade = $argumentos[5]
        $setor = $argumentos[6]
        $dispositivo = $argumentos[7]
        $patrimonio = $argumentos[8]
        $patriMonitor1 = $argumentos[9]
        $patriMonitor2 = $argumentos[10]
        $chamadoPar = $argumentos[11]
        $observacao = $argumentos[12]

        $argumentos | ForEach-Object { Write-Host $_ }
    
        $filePath = "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\ALTERACAO_KACE_${usuario}_${model}_${serviceTag}_${status}.txt"

        $conteudo = @"
TIPO: ${tipo}
SERVICE TAG: ${model}_${serviceTag}
USUARIO: ${usuario}
STATUS DO ATIVO: ${status}
LOCAL DO ATIVO: ${localidade}
DEPARTAMENTO: ${setor}
DISPOSITIVO: ${dispositivo}
PATRIMONIO: ${patrimonio} ${patriMonitor1} ${patriMonitor2}
CHAMADO PAR: ${chamadoPar}
OBSERVAÇÃO: ${observacao}
"@
    
        New-Item -Path $filePath -ItemType "File" -Value $conteudo -Force
        Start-Process "notepad.exe" $filePath 

    }

    if ($tipo -eq "DOCKSTATION")
    {
        $model = $argumentos[1]
        $serviceTag = $argumentos[2]
        $usuario = $argumentos[3]
        $status = $argumentos[4]
        $localidade = $argumentos[5]
        $setor = $argumentos[6]
        $dispositivo = $argumentos[7]
        $chamadoPar = $argumentos[8]
        $observacao = $argumentos[9]

        $argumentos | ForEach-Object { Write-Host $_ }
    
        $filePath = "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\ALTERACAO_KACE_${usuario}_${model}_${serviceTag}_${status}.txt"

        $conteudo = @"
TIPO: ${tipo}
SERVICE TAG: ${model}_${serviceTag}
USUARIO: ${usuario}
STATUS DO ATIVO: ${status}
LOCAL DO ATIVO: ${localidade}
DEPARTAMENTO: ${setor}
DISPOSITIVO: ${dispositivo}
CHAMADO PAR: ${chamadoPar}
OBSERVAÇÃO: ${observacao}
"@
    

        New-Item -Path $filePath -ItemType "File" -Value $conteudo -Force
        Start-Process "notepad.exe" $filePath
    }  
        
}

function Get-MonitorInfo($valor) {
    if (![string]::IsNullOrWhiteSpace($valor)) {
        $valor = $valor -replace ",", ""
        $partes = $valor -split " - "
        $partes = $partes | ForEach-Object { $_.Trim() }

        if ($partes.Count -ge 2) {
            return @($partes[0], $partes[1])
        }
    }
    return @("", "")  
}

function criadorTermoKace {
    param(
        [String[]] $dataTermo
    )

    $exePath = "C:\Users\PdfWriterApp\bin\Debug\net8.0\PdfWriterApp.exe"

    if (-not $dataTermo -or $dataTermo.Count -lt 11) {
        Write-Warning "Argumentos insuficientes: foram recebidos $($dataTermo.Count)."
        return
    }

    $quotedArgs = $dataTermo | ForEach-Object { '"{0}"' -f $_ }

    Write-Host "Argumentos sendo enviados para o app:"
    $quotedArgs | ForEach-Object { Write-Host $_ }

    try {
        Start-Process -FilePath $exePath -ArgumentList $quotedArgs -Wait
    }
    catch {
        Write-Host "Erro ao executar o processo:"
        Write-Host $_
        Write-Host $_.ScriptStackTrace
    }
}

function formAlteracaoKace
{
    param(
        [String]$tipo,
        [String]$model,
        [String]$serviceTag
    )

    $usuario = formInput -message "Informe o usuário (Ex.: filipi.serpa):"
    
    $UserAdmin = $env:USERNAME.Replace(".ope", "")

    if (isUserAd($usuario))
    {
        $entry = @("O usuário ${usuario} existe.", "Sucesso", 'Information')
        alertMessage($entry)
    } else 
    {
        $entry = @("O usuário ${usuario} não existe ou está incorreto.", "Exceção", 'Error')
        alertMessage($entry)
        return $null
    }

    $dataFormatada = Get-Date -Format "dd-MM-yyyy_HH-mm"

    New-Item -Path "C:\Users\master.user\Documentos\LOG_ABERTURA_CHAMADOS\ABERTURA_CHAMADO_${usuario}_${serviceTag}_${dataFormatada}.txt" -ItemType "file" -Value "
LOGIN: ${usuario}
SERVICE TAG: ${serviceTag}
DESC:
A/C: ${UserAdmin}" -Force

    $statusAtivo =  statusAtivo("")

    $localAtivo = localAtivo("")

    $dpAtivo = departamentoAtivo("")

    if ($tipo -eq "DESKTOP" -or $tipo -eq "NOTEBOOK" )
    {
        $haveMonitor = formConfirmAction ("Possui um monitor?", "Sim", "Nao")
        if ($haveMonitor -eq 'S') {
            $arrModeloMonitor = @()
            $arrServiceTagMonitor = @()

            getMonitores("") | ForEach-Object {
            
                for ($i = 0; $i -lt $monitores.Count; $i++)
                {
                 $modeloMon, $serviceTagMon = @($monitores[$i].Split("_"))
                 $arrModeloMonitor += $modeloMon
                 $arrServiceTagMonitor += $serviceTagMon
                }

                Write-Host $arrModeloMonitor $arrServiceTagMonitor
            }
        } else {
            $arrModeloMonitor = @()
            $arrServiceTagMonitor = @()
            $infoMonitores = ""
        }

        $hasDockstation = formConfirmAction("Possui Dockstation?", "Sim", "Não")
        if ($hasDockstation -eq 'S')
        {
            $dsAtivo = modeloDockAtivo
            $dsInfo = @($dsAtivo.split("_").toUpper())
            $dsData = $dsInfo[1] + "_" + $dsInfo[2]
            $data_dockstation = $dsInfo[1] + " / " + $dsInfo[2]
            Write-Host $dsInfo $data_dockstation
        } else 
        {
            $dsInfo = ""
        }

        $entryPatrimonio = getPatrimonioAtivos($tipo)

        if ($entryPatrimonio -and $entryPatrimonio.Count -ge 3) {
            $strPatrimonio = $entryPatrimonio[0]
            $patrimonio_comp = $entryPatrimonio[1]
            $arrMonPatri = $entryPatrimonio[2]
        } else {
            $strPatrimonio = ""
            $patrimonio_comp = ""
            $arrMonPatri = @()
        }

        Start-Process "notepad.exe" "C:\Users\master.user\Documentos\LOG_ABERTURA_CHAMADOS\ABERTURA_CHAMADO_${usuario}_${serviceTag}_${dataFormatada}.txt"
        
        $perifericos = perifericoAtivo("")

        $numChamado = formInput -message "Informe o chamado Par (Ex.: 177773): "
        
        $observacao = formInput -message "Tem alguma observação?"

        scriptTermo @(
        $tipo
        $model
        $serviceTag
        $usuario
        $statusAtivo
        $localAtivo
        $dpAtivo
        $infoMonitores[0]
        $infoMonitores[1]
        $dsData
        $strPatrimonio
        $arrMonPatri[0]
        $arrMonPatri[1]
        $numChamado
        $observacao
        )
        
        $data_dockstation = if ($data_dockstation) { $data_dockstation } else { "" }
        $patrimonio_comp = if ($patrimonio_comp) { $patrimonio_comp } else { "" }
        $entryPatrimonio = if ($entryPatrimonio) { $entryPatrimonio } else { @("") }
        $arrModelMonitor = if ($arrModelMonitor) { $arrModelMonitor } else { @("") }
        $arrServiceTagMonitor = if ($arrServiceTagMonitor) { $arrServiceTagMonitor } else { @("") }
        $arrMonPatri = if ($arrMonPatri) { $arrMonPatri } else { @("") }
        $perifericos = if ($perifericos) { $perifericos } else { @("") }

        
        $data_form = @(
            $tipo,
            $model,
            $serviceTag,
            $usuario,
            $statusAtivo,
            $localAtivo,
            $UserAdmin.Replace(".ope", ""),
            ($arrModeloMonitor -join ','),  
            ($arrServiceTagMonitor -join ','),
            ($arrMonPatri -join ','),
            $patrimonio_comp,
            $data_dockstation,
            ($perifericos -join ',')
        )

        criadorTermoKace -dataTermo $data_form

        return

    } elseif ($tipo -eq "MONITOR")
    {
        $hasDispositivo = formConfirmAction("Possui um dispositivo?", "Sim", "Nao")
        
        if($hasDispositivo -eq 'S')
        {
            $dados = getDadosDispositivos("");

            $tipoDisp = $dados[1]
            $modeloDisp = $dados[2]
            $serviceTagDisp = $dados[3]
            $infoDispositivo = $modeloDisp + "_" + $serviceTagDisp

        } else 
        {
            $infoDispositivo = " "
        }
        
        $entryPatrimonio = getPatrimonioAtivos($tipoDisp)

        $numChamado = formInput -message "Informe o chamado Par (Ex.: 177773): "
        
        $observacao = formInput -message "Tem alguma observação? "

        scriptTermo @(
        $tipo
        $model
        $serviceTag
        $usuario
        $statusAtivo
        $localAtivo
        $dpAtivo
        $infoDispositivo
        $entryPatrimonio
        $numChamado
        $observacao
        )
        

    } elseif ($tipo -eq "DOCKSTATION")
    {
        $hasDispositivo = formConfirmAction("Possui um dispositivo?", "Sim", "Nao")
        if($hasDispositivo -eq 'S')
        {
            $entryDisp = getDadosDispositivos("")
            $tipoDisp, $modelDisp, $svTagDisp = $entryDisp.split("_").toUpper()
            $dadosDisp = "${modelDisp}_${svTagDisp}"
        } else 
        {
            $tipoDisp = $null
            $dadosDisp = ""
        }

        $numChamado = formInput -message "Informe o chamado Par (Ex.: 177773):"
        
        $observacao = formInput -message "Tem alguma observação?"

        New-Item -Path "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\ALTERACAO_KACE_${usuario}_${model}_${serviceTag}_${statusAtivo}.txt" -ItemType "file" -Value "
TIPO: ${tipo}
SERVICE TAG: ${model}_${serviceTag}
USUARIO: ${usuario}
STATUS DO ATIVO: ${statusAtivo}
LOCAL DO ATIVO: ${localAtivo}
DEPARTAMENTO: ${dpAtivo}
DISPOSITIVO: ${dadosDisp}
CHAMADO PAR: ${numChamado}
OBSERVAÇÃO: ${observacao}" -Force
        
        Invoke-Item -Path "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\ALTERACAO_KACE_${usuario}_${model}_${serviceTag}_${statusAtivo}.txt"
        
        $entryAlert = @("O arquivo ALTERACAO_KACE_${usuario}_${model}_${serviceTag}_${statusAtivo} foi criado com sucesso.", "Sucesso", 'Information')
        alertMessage($entryAlert)
    }

    return
}

function formCriacaoKace
{
    param(
        [String]$tipo,
        [String]$model,
        [String]$serviceTag
    )
    $usuario = formInput -message "Informe o usuário (Ex.: filipi.serpa)"
    $userAd = Get-ADUser -Identity "${usuario}"

    $nfAtivo = formInput -message "Informe a Nota Fiscal do dispositivo"
    
    $numPedido = formInput -message "Informe o Nº Pedido do dispositivo"

    $statusAtivo =  statusAtivo("")

    $localAtivo = localAtivo("")

    $dpAtivo = departamentoAtivo("")

    $observacao = formInput -message "Tem alguma observação?"

    New-Item -Path "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\CRIACAO_KACE_${usuario}_${model}_${serviceTag}_${statusAtivo}.txt" -ItemType "file" -Value "
TIPO: ${tipo}
SERVICE TAG: ${model}_${serviceTag}
NOTA FISCAL: ${nfAtivo}
PEDIDO: ${numPedido}
USUARIO: ${usuario}
STATUS DO ATIVO: ${statusAtivo}
LOCAL DO ATIVO: ${localAtivo}
DEPARTAMENTO: ${dpAtivo}
OBSERVAÇÃO: ${observacao}" -Force
    Start-Process "notepad.exe" "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\CRIACAO_KACE_${usuario}_${model}_${serviceTag}_${statusAtivo}.txt"
   
    $entryAlert = @("O arquivo CRIACAO_KACE_${usuario}_${model}_${serviceTag}_${statusAtivo} foi criado com sucesso.", 'Information')
    alertMessage($entryAlert)

    return
}

function formReparoDell
{
    param(
        [String]$tipo,
        [String]$model,
        [String]$serviceTag
    )
    
    $usuario = formInput -message "Informe o usuário (Ex.: filipi.serpa)"
    $userAd = Get-ADUser -Identity "${usuario}"

    $localAtivo = localAtivo("")

    $dataIncidente = setValueByRegex("^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/\d{4}$", "Qual a data do incidente? (Ex.: 03/03/2025)")
    
    $descricaoProblema = setValueByRegex("[A-Z]{1,255}", "Descreva o problema apresentado: ")

    $depEquipamento = departamentoAtivo("")

    $confirmGarantia = formConfirmAction("O equipamento está na garantia?", "Sim", "Nao")

    if($valueConfirm -eq 'S') 
    {
        $garantia = "SIM";
    } else {
        $garantia = "NAO"
    }

    $estados = @(
                        "EM ANALISE"
                        "EM REPARO"
                        "FECHADO SEM REPARO"
                        "FECHADO"
                        "REPARO REALIZADO"
                )
    $status = Show-OptionForm -titulo "Selecione o STATUS da solicitação: " -optns $estados

    if ([String]::IsNullOrWhiteSpace($status))
    {
        $status = "EM ANALISE"
    }

    $confirmDataReparo = formConfirmAction("Já possui data prevista para reparo?", "Sim", "Nao")

    if ($confirmDataReparo -eq 'S')
    {
        $dataReparo = setValueByRegex("^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/\d{4}$", "Informe a data para reparo: (Ex.: 03/03/2025)")
    } else 
    {
        $dataReparo = "EM ANALISE"
    }

    $confirmValorReparo = formConfirmAction("Já possui o valor do reparo?", "Sim", "Nao")

    if($confirmValorReparo -eq 'S')
    {
        $regValor = setValueByRegex("\d{1,4},\d{2}", "Qual o valor do reparo? (Ex.: R$ 2500,00)")
        $valorReparo = "R$" + $regValor
    } else 
    {
        $valorReparo = "EM ANALISE"
    }

    $numChamado = formInput -message "Informe o chamado de participação: "

    $observacao = setValueByRegex("[A-Z]{1,255}", "Tem alguma observação?")

    New-Item -Path "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\REPARO_DELL_${usuario}_${model}_${serviceTag}.txt" -ItemType "file" -Value "

LOCAL: ${localAtivo}
TIPO DE EQUIPAMENTO: ${tipo}
SERVICE TAG: ${serviceTag}
DATA DO INCIDENTE: ${dataIncidente}
USUARIO: ${usuario}

SETOR: ${depEquipamento}
CAUSA DO ACIDENTE: ${descricaoProblema}

GARANTIA: ${garantia}
VALOR REPARO: ${valorReparo}
DATA REPARO: ${dataReparo}
STATUS: ${status}

CHAMADO PAR: ${numChamado}

OBSERVAÇÃO: ${observacao}" -Force

    Start-Process "notepad.exe" "C:\Users\master.user\Documentos\MOVIMENTACOES_KACE\REPARO_DELL_${usuario}_${model}_${serviceTag}.txt"

    $entryAlert = @("O arquivo REPARO_DELL_${usuario}_${model}_${serviceTag} foi criado com sucesso.", 'Information')
    alertMessage($entryAlert)
}

function createBtnAct {
    param (
        [ScriptBlock]$acao,
        [string]$nameBtn,
        [int[]]$dimensoes,
        $paramAct = $null
    )

    $entryBtn = @($nameBtn, $dimensoes[0], $dimensoes[1])
    $btn = btnGen -txt $nameBtn -px $dimensoes[0] -py $dimensoes[1]

    $btn.Size = New-Object System.Drawing.Size(140, 60)

    $localAcao = $acao
    $localParam = $paramAct

    $btn.Add_Click({
        & $localAcao @localParam
    }.GetNewClosure())

    return $btn
}

function createCheckAct {
    param (
        [string]$nameChk,
        [int[]]$dimensoes,
        $valorMarcado
    )

    $chk = checkBoxGen $nameChk

    $chk.Location = New-Object Drawing.Point($dimensoes[0], $dimensoes[1])

    $chk.Size = New-Object Drawing.Size(300, 40)

    $chk.Tag = $valorMarcado

    return $chk
}

function modeloDeskAtivo()
{
    $modelosDesktop = @(
        "DESKTOP_3000"
        "DESKTOP_3080"
        "DESKTOP_3090"
        "DESKTOP_3650"
        "Voltar"
    )
    $modeloDesk = gridSelectModelComputers -action ${function:operationFormKace} -models $modelosDesktop
    return $modeloDesk
}

function modeloNoteAtivo()
{
    $modelosNotebook = @(
        "NOTEBOOK_5420"
        "NOTEBOOK_5430"
        "NOTEBOOK_5440"
        "NOTEBOOK_5450"
        "NOTEBOOK_3570"
        "NOTEBOOK_3591"
        "NOTEBOOK_7420"
        "Voltar"
    )
    $modeloNote = gridSelectModelComputers -action ${function:operationFormKace} -models $modelosNotebook
    return $modeloNote
}

function modeloMonAtivo()
{
    $modelosMonitor = @(
        "MONITOR_E2222HS"
        "MONITOR_P2222HS"
        "MONITOR_P2722HS"
        "MONITOR_P2722H"
        "MONITOR_P2725H"
        "MONITOR_P2222H"
        "MONITOR_E2225H"
        "Voltar"
    )

   $modeloMonitor = gridSelectModel -action ${function:operationFormKace} -models $modelosMonitor
   return $modeloMonitor
}

function modeloDockAtivo()
{
    $modelosDockstation = @(
        "DOCKSTATION_WD19"
        "DOCKSTATION_WD19S"
        "DOCKSTATION_WD19C"
        "DOCKSTATION_WD19TB"
        "Voltar"
    )
    
    $modeloDockstation = gridSelectModel -action ${function:operationFormKace} -models $modelosDockstation
    return $modeloDockstation

}

function statusAtivo()
{
    $numStatusAtivo = Show-OptionForm -titulo "Qual o status do ativo?" -optns @("ATIVO", "RESERVADO")
    switch ($numStatusAtivo) {
        "ATIVO" { return "ATIVO" }
        "RESERVADO" { return "RESERVADO" }
        "Sair" { return $null}
    }
    return
}

function getDadosDispositivos()
{
    $dispositivo = Show-OptionForm -titulo "Qual o tipo de dispositivo?" -optns @("DESKTOP", "NOTEBOOK", "SAIR")
    switch ($dispositivo) {
        "DESKTOP" { return modeloDeskAtivo }
        "NOTEBOOK" { return modeloNoteAtivo }
        "Sair" { return }
    }
    return
}

function getPatrimonioAtivos($entry)
{
    $dispositivo = $entry

    if($dispositivo -ne $null)
    {
        $hasPatromonioPrincipal = formConfirmAction("O ${dispositivo} possui patrimonio?", "Sim", "Nao")
    }

    if($hasPatromonioPrincipal -eq 'S')
    {
        $compPatrimonio = formInput -message "Informe o nº do patrimônio do ${dispositivo}:"
    } else 
    {
        $compPatrimonio = $null
    }

    $hasPatromonioMonitores = formConfirmAction("O(s) monitor(es) possui(em) patrimonio?", "Sim", "Nao")
    
    if($hasPatromonioMonitores -eq 'S')
    {
        $monPatrimonio = formInput -message "Informe o nº do patrimônio do(s) monitor(es):"
        $arrMonPatri = $monPatrimonio.split(" ")
    } else 
    {
        $monPatrimonio = ""
    }
    
    if ($monPatrimonio -eq "" -and $compPatrimonio -eq "")
        {
            return @("", "", @())
        }
    if ($arrMonPatri.Count -eq 3)
        {
            $entryPatrimonio = "${compPatrimonio} (${dispositivo}) " + $arrMonPatri[0] + " (MONITOR) " + $arrMonPatri[1] + " (MONITOR) " + $arrMonPatri[2] + " (MONITOR) "
            return @($entryPatrimonio, $compPatrimonio, $arrMonPatri)
        }
    if ($arrMonPatri.Count -eq 2)
        {
            $entryPatrimonio = "${compPatrimonio} (${dispositivo}) " + $arrMonPatri[0] + " (MONITOR) " + $arrMonPatri[1] + " (MONITOR) "
            return @($entryPatrimonio, $compPatrimonio, $arrMonPatri)
        }
    if ($arrMonPatri.Count -eq 1)
        { 
            $entryPatrimonio = "${compPatrimonio} (${dispositivo}) " + $arrMonPatri[0] + " (MONITOR) "
            return @($entryPatrimonio, $compPatrimonio, $arrMonPatri) 
        }
    if ($compPatrimonio -ne $null)
        {
            $entryPatrimonio = "${compPatrimonio} (${dispositivo})"
            return @($entryPatrimonio, $compPatrimonio, @())
        }
}

function getMonitores()
{
    $opcoes = @(
        "Somente um"
        "2 Monitores"
        "3 Monitores"
        "Voltar"
    )
    $numMonitores = Show-OptionForm -titulo "Possui quantos monitores?" -optns $opcoes
    switch($numMonitores)
    {
        "Somente um" {  
                $monUm = modeloMonAtivo
                $arrMonitores = $monUm.replace("MONITOR_", "").split("_")
                return $arrMonitores[1].toUpper() + "_" + $arrMonitores[2].toUpper()
            }
        "2 Monitores" {
                $monUm = modeloMonAtivo("")
                $monDois = modeloMonAtivo("")
                $arrMonitores = $monUm.replace("MONITOR_", "").split("_") + $monDois.replace("MONITOR_", "").split("_")
                return $arrMonitores[1] + "_" + $arrMonitores[2].toUpper() + " " + $arrMonitores[4] + "_" + $arrMonitores[5].toUpper()
            }
        "3 Monitores" {
                $monUm = modeloMonAtivo("")
                $monDois = modeloMonAtivo("")
                $monTres = modeloMonAtivo("")
                $arrMonitores = $monUm.split("_").replace("MONITOR_", "") + $monDois.split("_").replace("MONITOR_", "") + $monTres.split("_").replace("MONITOR_", "")
                return $arrMonitores[1] + "_" + $arrMonitores[2].toUpper() + " " + $arrMonitores[4] + "_" + $arrMonitores[5].toUpper() + " " + $arrMonitores[7] + "_" + $arrMonitores[8].toUpper()
            }
        "Voltar" {
            return operationFormKace("")
        }
    }
    getMonitores("")
}

function getUserOrgUnit ($userAd)
{
    $dataUser = Get-ADUser -Identity $userAd

    if ($dataUser -eq "")
    {
        $entryAlert = @("Nenhum unidade organizacional encontrada. Contate o Supervisor Direto.", 'Error')
        alertMessage($entryAlert)
    } else 
    {
        $userOus = $dataUser -split ','
        $entryAlert = @($userOus, "OU", 'Information')
        alertMessage($entryAlert)
    }
}

function getUserGroups ($userAd)
{
     $userGroups = (Get-ADUser -Identity $userAd -Properties MemberOf).MemberOf | ForEach-Object {
    ($_ -split ',')[0] -replace 'CN='
    }

    if ($userGroups -eq "")
    {
        $entryAlert = @("Nenhum grupo encontrato. Contate o Supervisor Direto.", 'Error')
        alertMessage($entryAlert)
        return $null
    } else 
    {
        $entryAlert = @("Grupos do usuário ${userAd}:`n`n${userGroups}`n`n", "Grupos", 'Information')
        alertMessage($entryAlert)
        return $userGroups
    }
}

function setValueByRegex($entryValue)
{   
    $valueRegex, $msg = $entryValue.split(" ")
    $value = formInput -message $msg

    if ($value -Match $valueRegex)
    {
        return $value
    } else
    {
        setValueByRegex($valueRegex)
    }
    
}

function localAtivo()
{
    
    $locais = @(
        "SEDE CIDADE1"
        "SEDE CIDADE6"
        "LOJA CIDADE1"
        "LOJA CIDADE2"
        "LOJA CIDADE3"
        "LOJA CIDADE4"
        "ETA CIDADE1"
        "ETA CIDADE2"
        "ETA CIDADE1DIST2"
        "ETA CIDADE3"
        "ETA CIDADE4"
        "ALMOXARIFADO CIDADE1"
        "ALMOXARIFADO CIDADE2"
        "ALMOXARIFADO CIDADE3"
        "ALMOXARIFADO CIDADE4"
        "GSO CIDADE12"
        "GSO CIDADE4"
    )

    Show-OptionForm -titulo "Selecione a cidade desejada:" -optns $locais
}

function departamentoAtivo()
{
    $departamentos = @(
        "DEP1CDD1"
        "DEP2CDD4"
        "DEP3CDD1"
        "DEP4CDD2"
    )

    Show-OptionForm -titulo "Selecione o departamento:" -optns $departamentos
}

function perifericoAtivo()
{
    $optns = @{
       "KIT MOUSE TECLADO S/ FIO" = "KIT MOUSE TECLADO S/ FIO"
       "HEADPHONE"               = "HEADPHONE"
       "TECLADO"                 = "TECLADO"
       "MOUSE"                   = "MOUSE"
       "MOCHILA"                 = "MOCHILA"
    }

    $perifericos = Show-MultiOptionForm -titulo "Selecione os itens desejados:" -optns $optns
    return $perifericos
}

function checkSetInput($msgInput)
{
    do
    {
        $valorInput = getInput("`n${msgInput}")
    } while ($valorInput -eq $null -or $valorInput -eq "" -or $valorInput -eq " ")

    return $valorInput
}

function Show-OptionForm
{
    param(
        [string]$titulo = "Selecione uma opção:",
        [string[]]$optns
    )

    do
    {
        $selectOptn = [ref] ""

        $btnWidth = 140
        $btnHeight = 60
        $paddingX = 10
        $paddingY = 10
        $maxPorLinha = 4

        if ($optns.Count -le 4) {
            $formWidth = ($btnWidth + ($paddingX * 2)) + 20
            $formHeight = (($btnHeight + $paddingY) * $optns.Count) + 60
        }
        else 
        {
            $totalLinhas = [Math]::Ceiling($optns.Count / $maxPorLinha)
            $totalColunas = [Math]::Min($optns.Count, $maxPorLinha)

            $formWidth = ($totalColunas * ($btnWidth + $paddingX)) + 40
            $formHeight = ($totalLinhas * ($btnHeight + $paddingY)) + 60
        }

        $entry = @($titulo, $formWidth, $formHeight)

        $form = formGen($entry)
        
        for ($i = 0; $i -lt $optns.Count; $i++) 
        {
        
         $localOptn = $optns[$i]

         if ($optns.Count -le 4) {
            $posX = $paddingX
            $posY = $paddingY + ($i * ($btnHeight + $paddingY))

            $act = {
                param($optn, $refVar, $frm)
                $refVar.value = $optn
                $frm.Close()
               }

            $btn = createBtnAct -acao $act -nameBtn $localOptn -dimensoes @($posX, $posY) -paramAct @($localOptn, $selectOptn, $form)
            
            $form.Controls.Add($btn)
         }
         else {
            $linha = [Math]::Floor($i / $maxPorLinha)
            $coluna = $i % $maxPorLinha

            $posX = $paddingX + ($coluna * ($btnWidth + $paddingX))
            $posY = $paddingY + ($linha * ($btnHeight + $paddingY))

            $act = {
                param($optn, $refVar, $frm)
                $refVar.value = $optn
                $frm.Close()
               }

            $btn = createBtnAct -acao $act -nameBtn $localOptn -dimensoes @($posX, $posY) -paramAct @($localOptn, $selectOptn, $form)
            
            $form.Controls.Add($btn)
         }
        }

        $form.TopMost = $true
        $form.ShowDialog() | Out-Null

        $strOptn = "Deseja confirmar?
_
[_  " + $selectOptn.Value + "  _]"

        $confirm = formConfirmAction($strOptn, "Confirmar", "Não")

    } while ($confirm.toUpper() -ne 'S')

    return $selectOptn.Value
}

function Show-MultiOptionForm {
    param (
        [string]$titulo = "Marque as opções:",
        [hashtable]$optns 
    )

    do {
        $selectedValues = [System.Collections.ArrayList]::new()

        $chkWidth = 300
        $chkHeight = 40
        $paddingX = 10
        $paddingY = 10
        $maxPorLinha = 2  

        if (-not ($optns -is [hashtable])) {
            throw "O parâmetro -optns precisa ser um hashtable. Tipo recebido: $($optns.GetType().FullName)"
        }

        $optnKeys = @($optns.Keys)
        $total = $optnKeys.Count

        $totalLinhas = [math]::Ceiling($total / $maxPorLinha)
        $formWidth = ($maxPorLinha * ($chkWidth + $paddingX)) + 80
        $formHeight = ($totalLinhas * ($chkHeight + $paddingY)) + 140  

        $form = formGen @($titulo, $formWidth, $formHeight)

        $checkboxes = @()

        for ($i = 0; $i -lt $total; $i++) {
            $linha = [Math]::Floor($i / $maxPorLinha)
            $coluna = $i % $maxPorLinha

            $posX = $paddingX + ($coluna * ($chkWidth + $paddingX))
            $posY = $paddingY + ($linha * ($chkHeight + $paddingY))

            $text = $optnKeys[$i]
            $valor = $optns[$text]

            $chk = createCheckAct -nameChk $text -dimensoes @($posX, $posY) -valorMarcado $valor
            $checkboxes += $chk
            $form.Controls.Add($chk)
        }

        foreach ($chk in $checkboxes) {
            if ($chk.Checked) {
                Write-Host "Checkbox marcado: $($chk.Text) - Tag: $($chk.Tag)"
            }
        }

        $btnConfirm = btnGen -txt "Confirmar" -px ($formWidth - 140) -py ($formHeight - 80)
        $btnConfirm.Size = New-Object Drawing.Size(100, 35)


        $btnConfirm.Add_Click({
        $selectedValues.Clear()
        foreach ($chk in $checkboxes) {
            if ($chk.Checked) {
                Write-Host "Checkbox marcado: $($chk.Text) - Tag: $($chk.Tag)"
                $null = $selectedValues.Add($chk.Tag)
            }
        }
        $form.Close()
    })


        $form.Controls.Add($btnConfirm)
        $form.TopMost = $true
        $form.ShowDialog() | Out-Null

        $strOptn = "Deseja confirmar?
_
[_  " + ($selectedValues -join ", ") + "  _]"
        $confirm = formConfirmAction($strOptn, "Confirmar", "Não")

    } while ($confirm.ToUpper() -ne "S")

    return ,$selectedValues.ToArray()
}


function gridSelectModelComputers 
{
    param(
        [ScriptBlock]$action,
        [String[]]$models
    )

    $typeModel = $models[0].split("_").toUpper() | Select-Object -First 1

    $modelSelect = Show-OptionForm -titulo "Selecione o modelo do(a) ${typeModel}" -optns $models

    if($modelSelect -eq "Voltar") { return & $action ""}

    if (-not $modelSelect)
    {
        $entryAlert = @("Nenhuma opção válida selecionada.", "Exceção", 'Error')
        alertMessage($entryAlert)
        gridSelectModel ($models)
    }
    
    $tipo, $model = $modelSelect.replace("OK\s+", "").split("_")
    $serviceTag = formInput -message "Informe a SERVICE TAG do(a) ${typeModel}:"

    $entryDisp = @($typeModel, $serviceTag)
    if (isDispositivoAd($entryDisp))
    {
        $msg = "O dispositivo " + $serviceTag.toUpper() + " existe no domínio."
        $entryAlert = @($msg, "Sucesso", 'Information')
        alertMessage($entryAlert)
        return @($tipo, $model, $serviceTag.ToUpper())

    } else 
    {
        $entryAlert = @("O dispositivo ${serviceTag} não foi encontrado no domínio, confirme se o nome informado está correto.", "Exceção", 'Error')
        alertMessage($entryAlert)
        return gridSelectModel ($models)
    }
}

function gridSelectModel 
{
    param(
        [ScriptBlock]$action,
        [String[]]$models
    )

    $typeModel = $models[0].split("_").toUpper() | Select-Object -First 1

        $modelSelect = Show-OptionForm -titulo "Selecione o modelo do(a) ${typeModel}" -optns $models

        if ($modelSelect -eq "Voltar") { return & $action ""}

        if (-not $modelSelect)
        {
            gridSelectModel($models)
        }
    $tipo, $model = $modelSelect.replace("OK\s+", "").split("_")
    $serviceTag = formInput -message "Informe a SERVICE TAG do(a) ${typeModel}:"

    return @($tipo, $model, $serviceTag.toUpper())
}

function formPassword
{
    param(
        [String]$text,
        [Int]$width = 385,
        [Int]$heidth = 250
    )
    $form = formGen(@("Troca de Senha", $width, $heidth))

    $labelSenha = labelGen -msg $text -py 10
    $form.Controls.Add($labelSenha)

    $passwordBox = textBoxGen -py 50
    $passwordBox.UseSystemPasswordChar = $true
    $form.Controls.Add($passwordBox)

    $labelConfirmaSenha = labelGen -msg "Confirme sua senha:" -py 75
    $form.Controls.Add($labelConfirmaSenha)

    $confirmPasswordBox = textBoxGen -py 110
    $confirmPasswordBox.UseSystemPasswordChar = $true
    $form.Controls.Add($confirmPasswordBox)

    $checkBox = checkBoxGen -txt "Exibir Senha"
    $form.Controls.Add($checkBox)

    $submitButton = btnGen -txt "Enviar" -px 122 -py 180
    $form.Controls.Add($submitButton)

    $checkBox.Add_CheckedChanged({
        if ($checkBox.Checked)
        {
            $passwordBox.UseSystemPasswordChar = $false
            $confirmPasswordBox.UseSystemPasswordChar = $false
        } else
        {
            $passwordBox.UseSystemPasswordChar = $true
            $confirmPasswordBox.UseSystemPasswordChar = $true
        }
    })

    $submitButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $form.Topmost = $true

    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        
        $password = [string]$passwordBox.Text.Trim()
        $confirmPasswd = [string]$confirmPasswordBox.Text.Trim()

        if($password -eq $confirmPasswd)
        {
            if (($password -is [string]) -and -not [string]::IsNullOrWhiteSpace($password))
            {
                return $password
            } else 
            {
                $entryAlert = @("Nenhum valor válido informado.", "Informativo", 'Information')
                alertMessage($entryAlert)
                return $null
            }
        } else 
        {
                $entryAlert = @("As senhas informadas não coinscidem.", "Error", 'Error')
                alertMessage($entryAlert)
            return formPassword -text "Digite sua senha corretamente: "
        }

        
    } else 
    {
        Write-Error -Message "Nenhuma opção selecionada ou foi cancelado."
        return
    }

}

function formInput
{
    param(
        [String]$cabecalho = "RPA EMPRESA",
        [String]$message,
        [Int]$height = 385,
        [Int]$width = 250
    )
    $form = formGen(@($cabecalho, $height, $width))

    $label = labelGen -msg $message
    $form.Controls.Add($label)

    $textBox = textBoxGen -px 62 -py 60
    $form.Controls.Add($textBox)

    $checkBox = checkBoxGen("Confirmo que a informacao esta correta.")
    $form.Controls.Add($checkBox)

    $okButton = btnGen -txt "OK" -px 130 -py 100
    $okButton.Enabled = $false
    $form.Controls.Add($okButton)

    $checkBox.Add_CheckedChanged({
        $okButton.Enabled = $checkBox.Checked
    })

    $okButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $checkBox.Checked) {
        return $textBox.Text.Trim()
    } else {
        Write-Error -Message "Nenhuma opção selecionada ou foi cancelado."
        return $null
    }

}

function formConfirmAction ($entry)
{
    $msg, $nameBtnOne, $nameBtnTwo = $entry
    $entryForm = @(
        "RPA Rio+" 
        385
        250
    )
    $form = formGen($entryForm)

    $label = labelGen -msg "${msg}" 
    $form.Controls.Add($label)

    $confirmButton = btnGen -txt $nameBtnOne -px 85 -py 80
    $form.Controls.Add($confirmButton)

    $cancelButton = btnGen -txt $nameBtntWO -px 190 -py 80
    $form.Controls.Add($cancelButton)

    $confirmButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Tag = "confirmar"
        $form.Close()
    })

    $cancelButton.Add_Click({
        $form.Tag = "repetir"
        $form.Close()
    })

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return 'S'
    } else 
    {
        return 'N'
    }
}

function alertMessage ($entry)
{
    Add-Type -AssemblyName System.Windows.Forms
    $msg, $cabeçalho, $typeAlert = $entry
    [System.Windows.Forms.MessageBox]::Show($msg, $cabeçalho, 'OK', $typeAlert)
}

function formGen ($entry)
{
    Add-Type -AssemblyName System.Windows.Forms
    
    $txt, $x, $y = $entry

    $form = New-Object Windows.Forms.Form
    $form.Text = "${txt}"
    $form.Size = New-Object Drawing.Size($x,$y)
    $form.StartPosition = "CenterScreen"
    return $form
}

function btnGen
{
    param(
        [String]$txt,
        [Int]$px,
        [Int]$py
    )
    $btn = New-Object Windows.Forms.Button
    $btn.Text = $txt
    $btn.Location = New-Object Drawing.Point($px,$py)
    return $btn
}

function labelGen 
{
    param(
        [String]$msg,
        [Int]$px = 62,
        [Int]$py = 20
    )
    $label = New-Object Windows.Forms.Label
    $label.Text = "${msg}"
    $label.Location = New-Object Drawing.Point($px, $py)
    $label.Size = New-Object Drawing.Size(350, 40)
    return $label
}

function textBoxGen
{
    param(
        [Int]$px = 62,
        [Int]$py = 60
    )
    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Location = New-Object Drawing.Point($px, $py)
    $textBox.Size = New-Object Drawing.Size(220,20)
    return $textBox
}

function checkBoxGen($txt)
{
    $checkBox = New-Object Windows.Forms.CheckBox
    $checkBox.Text = "${txt}"
    $checkBox.Location = New-Object Drawing.Point(62, 130)
    $checkBox.Size = New-Object Drawing.Size(300, 40)
    return $checkBox
}

function confirmaInput($msg)
{
    $valorInput = checkSetInput($msg)
    $confirmaInput = checkSetInput("Você informou: ${valorInput}. Confirmar? (S/N)")
    if ($confirmaInput.toUpper() -ne "S")
    {
        confirmaInput($msg)
        
    } else
    {
        return $valorInput
    }
}

function alterarSenha ($user)
{
    Set-ADAccountPassword -Identity "${user}" -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Alterar1234" -Force) -ErrorAction Stop
    Set-ADUser -Identity "${user}" -ChangePasswordAtLogon: $true
    
    $entryAlert = @("A senha do usuário ${user} foi resetada com sucesso.", "Sucesso", 'Information')
    alertMessage($entryAlert)

    return
}

function alterarSenhaOffline
{
    param(
        [string]$user,
        [string]$pssw
    )
    try 
    {
        Set-ADAccountPassword -Identity "${user}" -Reset -NewPassword (ConvertTo-SecureString "${pssw}" -AsPlainText -Force) -ErrorAction Stop

        Set-ADUser -Identity "${user}" -ChangePasswordAtLogon: $false

        $entryAlert = @("A senha do usuário ${user} informado foi resetada com sucesso.", "Sucesso", 'Information')
        alertMessage($entryAlert)
    }
    catch
    {
        $msgError = "Exceção:" + $_ + " | " + "Mais detalhes no console."
        $entryAlert = @("${msgError}", "Temos uma exceção!", 'Information')
        alertMessage($entryAlert)
        Write-Host $_.ScriptStackTrace
        Write-Host $pssw $user
    }

    return
}

function getInput($question) 
{
    return Read-Host -Prompt "${question}" 
}

function addGroup
{
    param(
        [String]$group,
        [String]$user
    )
    Add-ADGroupMember -Identity $group -Members $user -Confirm:$false -ErrorAction Stop 
    
    $entryAlert = @("Usuário ${user} adicionado ao grupo com sucesso.", "Sucesso", 'Information')
    alertMessage($entryAlert)

    return
}

function rmGroup
{
    param(
        [String]$group,
        [String]$user
    )

    Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false -ErrorAction Stop 
    
    $entryAlert = @("Usuário ${user} removido do grupo com sucesso.", "Sucesso", 'Information')
    alertMessage($entryAlert)
    
    return
}

function encerrar ()
{
    $entry = @(
        "Estamos encerrando o sistema.", "Até mais", 'Information'
    )
    alertMessage($entry)
    $formApp.close()
}

function writeMsg ($msg) 
{
    return Write-Output "`n${msg}`n"
}

try 
{
    app -UserAdmin $UserAdmin 
}
catch
{
    Write-Host $_
    Write-Host $_.ScriptStackTrace
    app("")
}
