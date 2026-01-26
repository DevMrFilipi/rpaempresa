function form
{
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Cria o Formulário
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "RPA Rio+"
    $form.Size = New-Object System.Drawing.Size(300,250)
    $form.StartPosition = "CenterScreen"

    # Label Usuário
    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Text = "Usuário:"
    $userLabel.Location = New-Object System.Drawing.Point(20, 30)
    $userLabel.AutoSize = $true
    $form.Controls.Add($userLabel)

    # TextBox Usuário
    $userBox = New-Object System.Windows.Forms.TextBox
    $userBox.Location = New-Object System.Drawing.Point(90, 25)
    $userBox.Size = New-Object System.Drawing.Size(160, 30)  # Aumentei a altura
    $userBox.Font = New-Object System.Drawing.Font("Segoe UI", 10) # Fonte legível
    $form.Controls.Add($userBox)

    # Label Senha
    $passLabel = New-Object System.Windows.Forms.Label
    $passLabel.Text = "Senha:"
    $passLabel.Location = New-Object System.Drawing.Point(20, 70)
    $passLabel.AutoSize = $true
    $form.Controls.Add($passLabel)

    # TextBox Senha
    $passBox = New-Object System.Windows.Forms.TextBox
    $passBox.Location = New-Object System.Drawing.Point(90, 65)
    $passBox.Size = New-Object System.Drawing.Size(160, 30)
    $passBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $passBox.UseSystemPasswordChar = $true
    $form.Controls.Add($passBox)

    $showPassword = New-Object System.Windows.Forms.CheckBox
    $showPassword.Location = New-Object System.Drawing.Point(90, 100)
    $showPassword.Text = "Exibir senha"
    $form.Controls.Add($showPassword)

    # Botão OK
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(110, 130)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Add_Click({ 
    $form.Close()
    })
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $form.Topmost = $true

    $showPassword.Add_CheckedChanged({
        if ($showPassword.Checked)
        {
            $passBox.UseSystemPasswordChar = $false
        } else
        {
            $passBox.UseSystemPasswordChar = $true
        }
    })

    $result = $form.ShowDialog()

    # Mostra o formulário
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $userAdmin = $userBox.Text
        $userAdminDomain = "DOMINIO\" + $userAdmin
        $password = $passBox.Text | ConvertTo-SecureString -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential ($userAdminDomain, $password)
        
            Start-Process -FilePath powershell.exe -Credential $credential -ArgumentList @(
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-WindowStyle", "Hidden"
                "-File", "`"C:\Users\${userAdmin}\Documents\RPA_EMPRESA.ps1`"",
                "-UserAdmin", $userAdmin,
                "-NoExit"
            ) -Wait

        alertMessage -msg "Iniciado com sucesso" -cabecalho "Sucesso" -tipoAlerta 'Information'
    }
}

function alertMessage
{
    param(
        [string]$msg,
        [string]$cabecalho,
        [string]$tipoAlerta
    )
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($msg, $cabecalho, 'OK', $tipoAlerta)
}

try
    {
       form("") 
    }
    catch 
    {
        if ($_.Exception.Message -eq "Esse comando não pode ser executado devido ao erro: Nome de usuário ou senha incorretos.")
        {
            alertMessage -msg "Nome de usuário ou senha incorretos." -cabecalho "Falha no login" -tipoAlerta 'Error'
            form("")

        }

        Write-Host "Erro: $($_.Exception.Message)"
        Write-Host "StackTrace:"
        Write-Host $_.Exception.StackTrace
    }