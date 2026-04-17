$ErrorActionPreference = "Stop"

function Set-TextRangeStyle {
    param(
        [Parameter(Mandatory = $true)]$TextRange,
        [double]$FontSize = 18,
        [int]$Color = 0,
        [string]$FontName = "Aptos",
        [switch]$Bold
    )
    $TextRange.Font.Name = $FontName
    $TextRange.Font.Size = $FontSize
    $TextRange.Font.Color.RGB = $Color
    $TextRange.Font.Bold = [int]([bool]$Bold)
}

function Add-Textbox {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Text,
        [double]$FontSize = 18,
        [int]$Color = 0,
        [string]$FontName = "Aptos",
        [switch]$Bold,
        [int]$Align = 1
    )
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $shape.TextFrame2.WordWrap = $true
    $shape.TextFrame2.AutoSize = 0
    $shape.TextFrame2.VerticalAnchor = 3
    $shape.Line.Visible = 0
    $shape.Fill.Visible = 0
    $shape.TextFrame.TextRange.Text = $Text
    Set-TextRangeStyle -TextRange $shape.TextFrame.TextRange -FontSize $FontSize -Color $Color -FontName $FontName -Bold:$Bold
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Align
    return $shape
}

function Add-RoundedCard {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [int]$FillColor,
        [int]$LineColor = -1
    )
    $shape = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $shape.Fill.ForeColor.RGB = $FillColor
    if ($LineColor -ge 0) {
        $shape.Line.ForeColor.RGB = $LineColor
        $shape.Line.Weight = 1.25
    } else {
        $shape.Line.Visible = 0
    }
    return $shape
}

function Add-ImageFit {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [Parameter(Mandatory = $true)][string]$Path,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }
    $pic = $Slide.Shapes.AddPicture($Path, $false, $true, $Left, $Top, -1, -1)
    $ratio = [Math]::Min($Width / $pic.Width, $Height / $pic.Height)
    $pic.LockAspectRatio = -1
    $pic.Width = [Math]::Round($pic.Width * $ratio, 2)
    $pic.Height = [Math]::Round($pic.Height * $ratio, 2)
    $pic.Left = $Left + (($Width - $pic.Width) / 2)
    $pic.Top = $Top + (($Height - $pic.Height) / 2)
    return $pic
}

function Add-BulletList {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string[]]$Items,
        [double]$FontSize = 18,
        [int]$Color = 0,
        [string]$FontName = "Aptos"
    )
    $y = $Top
    $lineHeight = [Math]::Ceiling($FontSize + 7)
    foreach ($item in $Items) {
        $bullet = $Slide.Shapes.AddShape(9, $Left, $y + 4, 6, 6)
        $bullet.Fill.ForeColor.RGB = $Color
        $bullet.Line.Visible = 0
        Add-Textbox -Slide $Slide -Left ($Left + 14) -Top $y -Width ($Width - 14) -Height ($lineHeight + 4) -Text $item -FontSize $FontSize -Color $Color -FontName $FontName | Out-Null
        $y += $lineHeight
        if ($y - $Top -gt $Height) { break }
    }
}

function Add-TopBar {
    param(
        [Parameter(Mandatory = $true)]$Slide,
        [string]$Title,
        [string]$Subtitle = "",
        [string]$Badge = "",
        [hashtable]$Theme
    )
    Add-RoundedCard -Slide $Slide -Left 20 -Top 16 -Width 920 -Height 56 -FillColor $Theme.Navy | Out-Null
    Add-Textbox -Slide $Slide -Left 40 -Top 26 -Width 620 -Height 22 -Text $Title -FontSize 22 -Color $Theme.White -FontName $Theme.FontBold -Bold | Out-Null
    if ($Subtitle) {
        Add-Textbox -Slide $Slide -Left 40 -Top 48 -Width 620 -Height 14 -Text $Subtitle -FontSize 9.5 -Color $Theme.LightText -FontName $Theme.FontName | Out-Null
    }
    if ($Badge) {
        Add-RoundedCard -Slide $Slide -Left 782 -Top 28 -Width 120 -Height 28 -FillColor $Theme.Green | Out-Null
        Add-Textbox -Slide $Slide -Left 792 -Top 35 -Width 100 -Height 14 -Text $Badge -FontSize 10 -Color $Theme.White -FontName $Theme.FontBold -Bold -Align 2 | Out-Null
    }
}

$manualDir = "C:\Users\engenharia\Desktop\App LuisGEST - Cliente\Manual Operacao"
$shotsDir = Join-Path $manualDir "capturas_reais"
$pptxPath = Join-Path $manualDir "APRESENTACAO OPERACIONAL - LuisGEST.pptx"
$pdfPath = Join-Path $manualDir "APRESENTACAO OPERACIONAL - LuisGEST.pdf"

$theme = @{
    Navy      = 5515
    Blue      = 16735775
    Green     = 2021734
    White     = 16777215
    Bg        = 16447739
    Text      = 1710876
    Muted     = 7566195
    LightText = 16381173
    Border    = 15395562
    SoftBlue  = 16315637
    SoftGreen = 15596270
    SoftAmber = 15068893
    FontName  = "Aptos"
    FontBold  = "Aptos Display"
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$presentation = $ppt.Presentations.Add()
$presentation.PageSetup.SlideWidth = 960
$presentation.PageSetup.SlideHeight = 540

function New-Slide {
    param([int]$Index)
    return $presentation.Slides.Add($Index, 12)
}

# Slide 1 - cover
$slide = New-Slide 1
$slide.FollowMasterBackground = 0
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 120 -FillColor $theme.Navy | Out-Null
Add-Textbox -Slide $slide -Left 42 -Top 42 -Width 420 -Height 28 -Text "Manual Operacional" -FontSize 28 -Color $theme.White -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 42 -Top 78 -Width 420 -Height 28 -Text "luGEST" -FontSize 32 -Color $theme.White -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 480 -Top 44 -Width 380 -Height 20 -Text "Versao visual tipo apresentacao" -FontSize 16 -Color $theme.LightText -FontName $theme.FontName | Out-Null
Add-Textbox -Slide $slide -Left 480 -Top 72 -Width 380 -Height 16 -Text "Foco em Orcamentos, Encomendas, Planeamento e Operador" -FontSize 11 -Color $theme.LightText -FontName $theme.FontName | Out-Null
Add-RoundedCard -Slide $slide -Left 42 -Top 160 -Width 250 -Height 92 -FillColor $theme.SoftGreen -LineColor $theme.Border | Out-Null
Add-RoundedCard -Slide $slide -Left 312 -Top 160 -Width 250 -Height 92 -FillColor $theme.SoftBlue -LineColor $theme.Border | Out-Null
Add-RoundedCard -Slide $slide -Left 582 -Top 160 -Width 320 -Height 92 -FillColor $theme.SoftAmber -LineColor $theme.Border | Out-Null
Add-Textbox -Slide $slide -Left 58 -Top 182 -Width 220 -Height 16 -Text "Fluxo principal" -FontSize 11 -Color $theme.Muted -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 58 -Top 208 -Width 220 -Height 18 -Text "Orcamento -> Encomenda" -FontSize 18 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 328 -Top 182 -Width 220 -Height 16 -Text "Menus em destaque" -FontSize 11 -Color $theme.Muted -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 328 -Top 208 -Width 220 -Height 18 -Text "4 nucleos operacionais" -FontSize 18 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 598 -Top 182 -Width 280 -Height 16 -Text "Menus de apoio" -FontSize 11 -Color $theme.Muted -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 598 -Top 208 -Width 280 -Height 18 -Text "Materia-Prima, Produtos, Clientes" -FontSize 18 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 42 -Top 282 -Width 320 -Height 22 -Text "O que vais encontrar" -FontSize 20 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-BulletList -Slide $slide -Left 46 -Top 318 -Width 380 -Height 130 -Items @(
    "screenshots reais da aplicacao",
    "descricao simples dos menus principais",
    "exemplo pratico de fluxo completo",
    "formato visual mais proximo de uma apresentacao"
) -FontSize 13 -Color $theme.Muted -FontName $theme.FontName | Out-Null
Add-RoundedCard -Slide $slide -Left 466 -Top 282 -Width 438 -Height 196 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_orcamento_detalhe.png") -Left 472 -Top 288 -Width 426 -Height 184 | Out-Null

# Slide 2 - agenda
$slide = New-Slide 2
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Estrutura do Documento" -Subtitle "Visao rapida do conteudo principal" -Badge "Agenda" -Theme $theme
Add-Textbox -Slide $slide -Left 46 -Top 104 -Width 300 -Height 22 -Text "Menus prioritarios" -FontSize 20 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-BulletList -Slide $slide -Left 50 -Top 144 -Width 360 -Height 160 -Items @(
    "Orcamentos",
    "Encomendas",
    "Planeamento",
    "Operador"
) -FontSize 16 -Color $theme.Text -FontName $theme.FontName | Out-Null
Add-Textbox -Slide $slide -Left 470 -Top 104 -Width 260 -Height 22 -Text "Menus de apoio" -FontSize 20 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-BulletList -Slide $slide -Left 474 -Top 144 -Width 300 -Height 150 -Items @(
    "Materia-Prima",
    "Produtos",
    "Clientes",
    "Boas praticas"
) -FontSize 16 -Color $theme.Text -FontName $theme.FontName | Out-Null
Add-RoundedCard -Slide $slide -Left 46 -Top 336 -Width 848 -Height 110 -FillColor $theme.SoftBlue -LineColor $theme.Border | Out-Null
Add-Textbox -Slide $slide -Left 64 -Top 362 -Width 280 -Height 16 -Text "Exemplo usado no manual" -FontSize 11 -Color $theme.Muted -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 64 -Top 388 -Width 420 -Height 18 -Text "ORC-2026-0001 -> BARCELBAL0001" -FontSize 20 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 64 -Top 418 -Width 760 -Height 16 -Text "Criado na base para demonstrar o fluxo completo: aprovacao, conversao, planeamento e operador." -FontSize 11 -Color $theme.Muted -FontName $theme.FontName | Out-Null

# Slide 3 - orcamentos
$slide = New-Slide 3
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Orcamentos" -Subtitle "Ponto de partida comercial e tecnico" -Badge "1 | Orcamentos" -Theme $theme
Add-Textbox -Slide $slide -Left 44 -Top 102 -Width 300 -Height 20 -Text "O que se faz aqui" -FontSize 19 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-BulletList -Slide $slide -Left 48 -Top 136 -Width 330 -Height 180 -Items @(
    "criar propostas e controlar estados",
    "adicionar linhas, DXF/DWG e artigos de stock",
    "configurar operacoes, nesting e desconto",
    "gerar PDF e converter em encomenda"
) -FontSize 13 -Color $theme.Muted -FontName $theme.FontName | Out-Null
Add-RoundedCard -Slide $slide -Left 44 -Top 360 -Width 330 -Height 110 -FillColor $theme.SoftGreen -LineColor $theme.Border | Out-Null
Add-Textbox -Slide $slide -Left 60 -Top 382 -Width 250 -Height 16 -Text "Leitura rapida" -FontSize 11 -Color $theme.Muted -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 60 -Top 408 -Width 290 -Height 40 -Text "O orcamento e o ponto onde nasce a proposta e onde se decide o que vai seguir para producao." -FontSize 11 -Color $theme.Text -FontName $theme.FontName | Out-Null
Add-RoundedCard -Slide $slide -Left 404 -Top 102 -Width 490 -Height 370 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_orcamento_detalhe.png") -Left 412 -Top 110 -Width 474 -Height 354 | Out-Null

# Slide 4 - orcamentos split
$slide = New-Slide 4
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Orcamentos" -Subtitle "Da lista ao detalhe" -Badge "1 | Orcamentos" -Theme $theme
Add-Textbox -Slide $slide -Left 42 -Top 102 -Width 180 -Height 18 -Text "Lista de orcamentos" -FontSize 16 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-RoundedCard -Slide $slide -Left 40 -Top 126 -Width 406 -Height 180 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_orcamentos_lista.png") -Left 46 -Top 132 -Width 394 -Height 168 | Out-Null
Add-BulletList -Slide $slide -Left 46 -Top 322 -Width 392 -Height 90 -Items @(
    "pesquisa por numero, cliente e estado",
    "abertura rapida do registo certo",
    "leitura simples do historico"
) -FontSize 11.5 -Color $theme.Muted -FontName $theme.FontName | Out-Null
Add-Textbox -Slide $slide -Left 476 -Top 102 -Width 180 -Height 18 -Text "Detalhe do orcamento" -FontSize 16 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-RoundedCard -Slide $slide -Left 474 -Top 126 -Width 420 -Height 180 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_orcamento_detalhe.png") -Left 480 -Top 132 -Width 408 -Height 168 | Out-Null
Add-BulletList -Slide $slide -Left 480 -Top 322 -Width 390 -Height 110 -Items @(
    "cliente, linhas, resumo financeiro e PDF",
    "operacoes, transporte e desconto",
    "botao de conversao em encomenda"
) -FontSize 11.5 -Color $theme.Muted -FontName $theme.FontName | Out-Null

# Slide 5 - encomendas
$slide = New-Slide 5
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Encomendas" -Subtitle "Trabalho real criado a partir do aprovado" -Badge "2 | Encomendas" -Theme $theme
Add-Textbox -Slide $slide -Left 44 -Top 102 -Width 300 -Height 18 -Text "Como ler este menu" -FontSize 19 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-BulletList -Slide $slide -Left 48 -Top 136 -Width 330 -Height 170 -Items @(
    "recebe pecas, materiais e tempos do orcamento",
    "organiza por material, espessura e peca",
    "mostra operacoes, estados e montagem",
    "liga a encomenda ao resto do sistema"
) -FontSize 13 -Color $theme.Muted -FontName $theme.FontName | Out-Null
Add-RoundedCard -Slide $slide -Left 404 -Top 102 -Width 490 -Height 370 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_encomenda_detalhe.png") -Left 412 -Top 110 -Width 474 -Height 354 | Out-Null

# Slide 6 - planeamento
$slide = New-Slide 6
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Planeamento" -Subtitle "Quadro semanal para distribuir carga" -Badge "3 | Planeamento" -Theme $theme
Add-RoundedCard -Slide $slide -Left 40 -Top 104 -Width 852 -Height 274 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_planeamento.png") -Left 48 -Top 112 -Width 836 -Height 258 | Out-Null
Add-BulletList -Slide $slide -Left 54 -Top 396 -Width 800 -Height 90 -Items @(
    "o backlog aparece no lado esquerdo por operacao",
    "o quadro semanal mostra os blocos colocados na semana",
    "os indicadores de carga ajudam a perceber ocupacao e historico"
) -FontSize 12.5 -Color $theme.Muted -FontName $theme.FontName | Out-Null

# Slide 7 - operador
$slide = New-Slide 7
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Operador" -Subtitle "Execucao operacional das pecas" -Badge "4 | Operador" -Theme $theme
Add-RoundedCard -Slide $slide -Left 40 -Top 104 -Width 852 -Height 274 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_operador_detalhe.png") -Left 48 -Top 112 -Width 836 -Height 258 | Out-Null
Add-BulletList -Slide $slide -Left 54 -Top 396 -Width 800 -Height 92 -Items @(
    "o operador escolhe a peca e a operacao certa",
    "pode iniciar, finalizar, interromper, dar baixa ou ver desenho",
    "a tabela mostra o que esta em curso e o que continua pendente"
) -FontSize 12.5 -Color $theme.Muted -FontName $theme.FontName | Out-Null

# Slide 8 - materia prima
$slide = New-Slide 8
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Materia-Prima" -Subtitle "Base de stock para reserva, nesting e consumo" -Badge "Apoio | Stock" -Theme $theme
Add-RoundedCard -Slide $slide -Left 40 -Top 104 -Width 852 -Height 274 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_materia_prima.png") -Left 48 -Top 112 -Width 836 -Height 258 | Out-Null
Add-BulletList -Slide $slide -Left 54 -Top 396 -Width 800 -Height 82 -Items @(
    "gerir lotes, chapas, retalhos e disponibilidade",
    "dar suporte direto ao nesting, reservas e compras",
    "consultar stock real por formato, espessura e localizacao"
) -FontSize 12.5 -Color $theme.Muted -FontName $theme.FontName | Out-Null

# Slide 9 - produtos/clientes
$slide = New-Slide 9
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Menus de Apoio" -Subtitle "Produtos e Clientes" -Badge "Apoio | Cadastros" -Theme $theme
Add-Textbox -Slide $slide -Left 42 -Top 102 -Width 160 -Height 18 -Text "Produtos" -FontSize 16 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-RoundedCard -Slide $slide -Left 40 -Top 126 -Width 406 -Height 180 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_produtos.png") -Left 46 -Top 132 -Width 394 -Height 168 | Out-Null
Add-BulletList -Slide $slide -Left 46 -Top 322 -Width 392 -Height 88 -Items @(
    "artigos internos e consumiveis",
    "base para montagem e stock auxiliar",
    "controlo de quantidades e unidades"
) -FontSize 11.5 -Color $theme.Muted -FontName $theme.FontName | Out-Null
Add-Textbox -Slide $slide -Left 476 -Top 102 -Width 160 -Height 18 -Text "Clientes" -FontSize 16 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-RoundedCard -Slide $slide -Left 474 -Top 126 -Width 420 -Height 180 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_clientes.png") -Left 480 -Top 132 -Width 408 -Height 168 | Out-Null
Add-BulletList -Slide $slide -Left 480 -Top 322 -Width 390 -Height 88 -Items @(
    "ficha comercial e fiscal do cliente",
    "moradas, contactos e email",
    "base para orcamentos e encomendas"
) -FontSize 11.5 -Color $theme.Muted -FontName $theme.FontName | Out-Null

# Slide 10 - flow
$slide = New-Slide 10
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Fluxo do Exemplo" -Subtitle "Do orcamento aprovado ate ao operador" -Badge "Exemplo real" -Theme $theme
$steps = @(
    @{x=52;  n='1'; t='Criar e guardar'; s='orcamento'},
    @{x=220; n='2'; t='Aprovar'; s='orcamento'},
    @{x=388; n='3'; t='Converter'; s='em encomenda'},
    @{x=556; n='4'; t='Planear'; s='corte laser'},
    @{x=724; n='5'; t='Iniciar'; s='peca no operador'}
)
foreach ($step in $steps) {
    Add-RoundedCard -Slide $slide -Left $step.x -Top 226 -Width 128 -Height 96 -FillColor $theme.White -LineColor $theme.Border | Out-Null
    $circle = $slide.Shapes.AddShape(9, $step.x + 14, 242, 28, 28)
    $circle.Fill.ForeColor.RGB = $theme.Blue
    $circle.Line.Visible = 0
    Add-Textbox -Slide $slide -Left ($step.x + 19) -Top 248 -Width 18 -Height 14 -Text $step.n -FontSize 11 -Color $theme.White -FontName $theme.FontBold -Bold -Align 2 | Out-Null
    Add-Textbox -Slide $slide -Left ($step.x + 52) -Top 246 -Width 64 -Height 16 -Text $step.t -FontSize 13 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
    Add-Textbox -Slide $slide -Left ($step.x + 52) -Top 268 -Width 64 -Height 28 -Text $step.s -FontSize 11 -Color $theme.Muted -FontName $theme.FontName | Out-Null
}
for ($i=0; $i -lt ($steps.Count - 1); $i++) {
    $line = $slide.Shapes.AddLine($steps[$i].x + 128, 274, $steps[$i+1].x, 274)
    $line.Line.ForeColor.RGB = $theme.Border
    $line.Line.Weight = 2
}
Add-RoundedCard -Slide $slide -Left 48 -Top 368 -Width 844 -Height 98 -FillColor $theme.SoftGreen -LineColor $theme.Border | Out-Null
Add-Textbox -Slide $slide -Left 64 -Top 390 -Width 220 -Height 16 -Text "Resultado do exemplo" -FontSize 11 -Color $theme.Muted -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 64 -Top 416 -Width 700 -Height 20 -Text "1 orcamento -> 1 encomenda -> 1 bloco planeado -> 1 peca iniciada" -FontSize 19 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null

# Slide 11 - close
$slide = New-Slide 11
$slide.Background.Fill.ForeColor.RGB = $theme.Bg
Add-RoundedCard -Slide $slide -Left 20 -Top 16 -Width 920 -Height 508 -FillColor $theme.White | Out-Null
Add-TopBar -Slide $slide -Title "Conclusao" -Subtitle "Como ler o luGEST de forma simples" -Badge "Fecho" -Theme $theme
Add-Textbox -Slide $slide -Left 44 -Top 112 -Width 410 -Height 22 -Text "Os 4 menus que mais importam no dia a dia" -FontSize 21 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-BulletList -Slide $slide -Left 48 -Top 154 -Width 410 -Height 160 -Items @(
    "Orcamentos: preparar e validar a proposta",
    "Encomendas: transformar o aprovado em trabalho real",
    "Planeamento: encaixar a carga na semana certa",
    "Operador: executar a peca com controlo operacional"
) -FontSize 14 -Color $theme.Text -FontName $theme.FontName | Out-Null
Add-RoundedCard -Slide $slide -Left 44 -Top 344 -Width 390 -Height 102 -FillColor $theme.SoftBlue -LineColor $theme.Border | Out-Null
Add-Textbox -Slide $slide -Left 60 -Top 366 -Width 170 -Height 16 -Text "Regra simples" -FontSize 11 -Color $theme.Muted -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 60 -Top 392 -Width 320 -Height 18 -Text "Comercial -> Producao -> Execucao" -FontSize 20 -Color $theme.Text -FontName $theme.FontBold -Bold | Out-Null
Add-Textbox -Slide $slide -Left 60 -Top 420 -Width 340 -Height 16 -Text "Se estes 4 menus estiverem dominados, o nucleo do software fica dominado." -FontSize 10.5 -Color $theme.Muted -FontName $theme.FontName | Out-Null
Add-RoundedCard -Slide $slide -Left 490 -Top 120 -Width 392 -Height 306 -FillColor $theme.White -LineColor $theme.Border | Out-Null
Add-ImageFit -Slide $slide -Path (Join-Path $shotsDir "manual_operador_detalhe.png") -Left 498 -Top 128 -Width 376 -Height 290 | Out-Null

$presentation.SaveAs($pptxPath)
$presentation.SaveAs($pdfPath, 32)
$presentation.Close()
$ppt.Quit()
Write-Output $pptxPath
Write-Output $pdfPath
