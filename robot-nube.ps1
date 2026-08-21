# ============================================================================
# ROBOT ACTUALIZADOR (NUBE) - Panel de Fleteros PPP
# Corre en GitHub Actions (ver .github/workflows/actualizar-panel.yml).
#   1. API de Gescom               -> efectividad de entrega (repartos), motivos,
#      estadisticas, analisis de rechazos, proveedores y feriados
#   2. Planilla de carton de Drive -> se descarga sola desde CARTON_URL (secreto)
# Genera data.js e historial-meses.json en la raiz del repo; el workflow los
# commitea y GitHub Pages publica. Credenciales por variables de entorno:
#   GESCOM_REALM / GESCOM_USUARIO / GESCOM_CLAVE / CARTON_URL (secretos del repo)
# Para probarlo en la PC: sin esas variables, lee robot\gescom-api.txt y el
# Carton*.xlsx de Documentos\GESCOM (modo prueba local).
# ============================================================================

$ErrorActionPreference = "Stop"

# --- Rutas ------------------------------------------------------------------
$RAIZ = $env:GITHUB_WORKSPACE
$MODO = "NUBE"
if (-not $RAIZ) {
  $RAIZ = Split-Path $MyInvocation.MyCommand.Path
  $MODO = "PRUEBA LOCAL"
}
$DIAS_HISTORIAL = 50   # margen de lectura hacia atras (cubre mes completo + prom. 14d)
$DIAS_PUBLICAR = 14    # fechas con datos que se publican en la web

# Filas que no van a la web (no-fleteros y excluidos a pedido de Lucas)
$EXCLUIR = @("SIN CHOFER", "RETIRA EN DEPOSITO",
             "LEANDRO BENITEZ", "MARCELO VACA", "GONZALO CALO", "EZEQUIEL HEREDIA",
             "CARLOS GUILLERMO ESCUDERO", "GABRIEL MAYMO", "SEMI JORGE")

function Log($msg) {
  Write-Output ((Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $msg)
}

# --- Lector generico de .xlsx (sin Excel) -----------------------------------
function Abrir-Xlsx($ruta) {
  $tmp = Join-Path $env:TEMP ("xlsx_" + [Guid]::NewGuid().ToString("N"))
  $zip = "$tmp.zip"
  Copy-Item $ruta $zip -Force
  Expand-Archive $zip $tmp -Force
  Remove-Item $zip -Force
  $strings = New-Object System.Collections.ArrayList
  $ssPath = Join-Path $tmp "xl\sharedStrings.xml"
  if (Test-Path $ssPath) {
    [xml]$ss = Get-Content $ssPath -Encoding UTF8
    foreach ($si in $ss.sst.si) {
      if ($si.t -is [System.Xml.XmlElement]) { [void]$strings.Add($si.t.InnerText) }
      elseif ($null -ne $si.t) { [void]$strings.Add([string]$si.t) }
      else { [void]$strings.Add((($si.r | ForEach-Object { $_.t.InnerText }) -join "")) }
    }
  }
  # nombre de hoja -> archivo xml
  [xml]$wb = Get-Content (Join-Path $tmp "xl\workbook.xml") -Encoding UTF8
  [xml]$rels = Get-Content (Join-Path $tmp "xl\_rels\workbook.xml.rels") -Encoding UTF8
  $mapaRel = @{}
  foreach ($r in $rels.Relationships.Relationship) { $mapaRel[$r.Id] = $r.Target }
  $hojas = [ordered]@{}
  foreach ($s in $wb.workbook.sheets.sheet) {
    $rid = $s.GetAttribute("r:id")
    if ($mapaRel[$rid]) { $hojas[$s.name] = Join-Path $tmp ("xl\" + $mapaRel[$rid].Replace("/", "\")) }
  }
  return @{ carpeta = $tmp; strings = $strings; hojas = $hojas }
}

function Leer-Hoja($libro, $nombreHoja) {
  # Devuelve lista de filas; cada fila es un hashtable Columna(letra) -> valor
  $ruta = $libro.hojas[$nombreHoja]
  if (-not $ruta) { return @() }
  [xml]$sh = Get-Content $ruta -Encoding UTF8
  $filas = New-Object System.Collections.ArrayList
  foreach ($row in $sh.worksheet.sheetData.row) {
    $cells = @{}
    foreach ($c in $row.c) {
      if ($c.r -match "^([A-Z]+)\d+$") { $col = $Matches[1] } else { continue }
      $v = $c.v
      if ($c.t -eq "s" -and $null -ne $v) { $v = $libro.strings[[int]$v] }
      $cells[$col] = $v
    }
    if ($cells.Count -gt 0) { [void]$filas.Add($cells) }
  }
  return $filas
}

function Cerrar-Xlsx($libro) {
  Remove-Item -Recurse -Force $libro.carpeta -ErrorAction SilentlyContinue
}

function EsNumero($v) { return ($null -ne $v -and [string]$v -match "^-?\d+(\.\d+)?$") }

# Porcentaje SIN decimales (regla de la empresa, 20/8): de ,50 para arriba
# redondea para arriba y de ,49 para abajo. Los premios se calculan sobre este
# numero, asi el fletero cobra por el mismo % que ve en el panel.
# OJO: [math]::Round por defecto redondea "al par" (88,5 -> 88 y 94,5 -> 94),
# que NO es lo que se pidio -> hay que forzar AwayFromZero.
function PctEntero($num, $den) {
  if ($den -le 0) { return $null }
  return [int][math]::Round(100.0 * $num / $den, 0, [System.MidpointRounding]::AwayFromZero)
}

# FORCE_MES=anterior -> resolver al mes CALENDARIO anterior (verificacion mensual
# de cierre; asi el mismo workflow sirve para cualquier mes sin tocar nada).
if ($env:FORCE_MES -eq "anterior") {
  $env:FORCE_MES = (Get-Date -Day 1).Date.AddMonths(-1).ToString("yyyy-MM")
}

# ============================================================================
Log ("================ INICIO (" + $MODO + ") ================")
if ($env:FORCE_MES) { Log ("Modo VERIFICACION del mes " + $env:FORCE_MES) }

# --- Bajar los datos UNA sola vez por dia ---------------------------------
# Hay varias corridas programadas por dia (respaldos por si GitHub demora o
# saltea alguna). Para no cargar la API de Gescom de mas, si el panel YA se
# actualizo hoy y esta es una corrida automatica, salimos sin bajar nada.
# Las corridas MANUALES (Run workflow) y las de VERIFICACION (FORCE_MES) siempre corren.
$dataActual = Join-Path $RAIZ "data.js"
if ($env:GITHUB_EVENT_NAME -eq "schedule" -and -not $env:FORCE_MES -and (Test-Path $dataActual)) {
  $cabecera = (Get-Content $dataActual -TotalCount 3 -Encoding UTF8) -join "`n"
  $hoyStr = (Get-Date -Format "yyyy-MM-dd")
  if ($cabecera -match ("Ultima actualizacion: " + [regex]::Escape($hoyStr))) {
    Log ("Ya actualizado hoy (" + $hoyStr + "): no hace falta bajar de nuevo. FIN.")
    exit 0
  }
}

# ============================================================================
# 2b) MOTIVOS, ESTADISTICAS Y ANALISIS DE RECHAZOS desde la API de Gescom
#     Desde jul-2026 reemplaza al CSV de ventas (verificado boleta por boleta:
#     el ImporteItem del CSV = importeTotal del item de la API, con IVA).
#     Si la API falla, el robot ABORTA sin publicar: el panel queda como ayer.
# ============================================================================
$motivos = @{}
$motivosPorChofer = @{}
$statsChofer = @{}
$anZonas = @(); $anVend = @(); $anClientes = @(); $anImporte = 0; $mesFE = ""
$anProveedores = @(); $anFacturado = 0
$choProvFact = @{}; $choProvRech = @{}; $provFact = @{}; $provRech = @{}; $factTotal = @{}
$impRechCho = @{}
$feriadosWeb = @()

# --- Conexion: credenciales por variables de entorno (secretos del repo) ---
# Trim: un salto de linea colado en un secreto rompe la URL del login (500)
$credApi = @{ REALM = ([string]$env:GESCOM_REALM).Trim(); USUARIO = ([string]$env:GESCOM_USUARIO).Trim(); CLAVE = ([string]$env:GESCOM_CLAVE).Trim() }
if (-not $credApi.REALM) {
  # Modo prueba local: leerlas del archivo de siempre
  $archCred = "C:\Users\luqaa\Documents\PPP-Fleteros\robot\gescom-api.txt"
  foreach ($lg in Get-Content $archCred -Encoding UTF8) {
    $pg = $lg.Split("=", 2); if ($pg.Count -eq 2) { $credApi[$pg[0].Trim()] = $pg[1].Trim() }
  }
}
if (-not $credApi.REALM -or -not $credApi.USUARIO -or -not $credApi.CLAVE) {
  Log "ERROR: faltan los secretos GESCOM_REALM / GESCOM_USUARIO / GESCOM_CLAVE"
  exit 1
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-TokenGescom {
  # El token de ACCESO dura pocos minutos (no 24 h): se renueva cada vez que hace falta
  $tk = (Invoke-RestMethod -Method Post -TimeoutSec 30 `
    -Uri ("https://auth.gescom.online/realms/" + $credApi.REALM + "/protocol/openid-connect/token") `
    -Body @{ grant_type = "password"; client_id = "gcw-web-api"; username = $credApi.USUARIO; password = $credApi.CLAVE }).access_token
  $script:HDR_GESCOM = @{ Authorization = "Bearer " + $tk }
}
try {
  Get-TokenGescom
} catch {
  Log ("ERROR: no pude entrar a la API de Gescom (token): " + $_.Exception.Message)
  Log "No se publica nada: el panel queda como estaba. Probar de nuevo mas tarde con el acceso directo."
  exit 1
}

function Get-Gescom($ruta) {
  # Reintenta con espera creciente (el servidor rebota rafagas: "Acceso denegado")
  # y renueva el token si vencio (401 No autorizado).
  $esperasG = @(10, 30, 60, 120, 180); $ng = 0
  while ($true) {
    try { return Invoke-RestMethod -Uri ("https://pehuenia.gescom.online/data/cmd/" + $ruta) -Headers $script:HDR_GESCOM -TimeoutSec 180 }
    catch {
      if ($ng -ge $esperasG.Count) { throw }
      $st = 0
      try { $st = [int]$_.Exception.Response.StatusCode } catch { }
      if ($st -eq 401) {
        try { Get-TokenGescom } catch { Start-Sleep -Seconds $esperasG[$ng] }
      } else {
        Start-Sleep -Seconds $esperasG[$ng]
      }
      $ng++
    }
  }
}

try {
  # --- Tablas de nombres (codigo -> nombre) ---
  # OJO: guardar la respuesta en una variable ANTES de recorrerla. Con
  # @(Get-Gescom ...) directo, PowerShell 5.1 envuelve el array JSON en un
  # array de 1 elemento y el foreach recorre "una sola cosa" (bug ya sufrido).
  $resp = Get-Gescom "ventas/api/v1/get-empleados?tipo=CHF"
  $nomChofer = @{}
  foreach ($x in @($resp)) {
    $nomChofer[[string]$x.codigo] = (([string]$x.nombre).Trim().ToUpper() -replace "\s+", " ")
  }
  $resp = Get-Gescom "ventas/api/v1/get-vendedores"
  $nomVend = @{}
  foreach ($x in @($resp)) {
    $nomVend[[string]$x.codigo] = (([string]$x.nombre).Trim().ToUpper() -replace "\s+", " ")
  }
  $resp = Get-Gescom "compras/api/v1/get-proveedores"
  $nomProv = @{}
  foreach ($x in @($resp)) {
    $nomProv[[string]$x.codigo] = ([string]$x.nombre).Trim()
  }
  $resp = Get-Gescom "ventas/api/v1/get-clientes"
  $cliLoc = @{}; $cliRaz = @{}
  foreach ($x in @($resp)) {
    $cliLoc[[string]$x.codigo] = ([string]$x.localidad).Trim().ToUpper()
    $cliRaz[[string]$x.codigo] = ([string]$x.razonSocial).Trim()
  }
  $resp = Get-Gescom "inventario/api/v2/get-articulos"
  $provArt = @{}
  foreach ($x in @($resp)) {
    $provArt[[string]$x.codigo] = [string]$x.codigoProveedor
  }
  $resp = Get-Gescom "ventas/api/v1/get-feriados"
  $feriadosWeb = @(@($resp) | ForEach-Object { [string]$_.fecha } | Sort-Object -Unique)
  Log ("API Gescom OK: " + $nomChofer.Count + " choferes, " + $nomVend.Count + " vendedores, " +
    $cliLoc.Count + " clientes, " + $provArt.Count + " articulos, " + $feriadosWeb.Count + " feriados")

  # --- Ventas dia por dia (fechadesde/fechahasta filtran por fecha de CARGA,
  #     y la preventa se carga hasta ~3 semanas antes de la entrega: margen 21 dias.
  #     OJO: fechahasta es EXCLUSIVA) ---
  $hoyDt = (Get-Date).Date
  $mesIniDt = (Get-Date -Day 1).Date
  # Escape hatch: FORCE_MES=yyyy-MM recalcula un mes YA CERRADO (backfill del
  # historial). Fuerza la ventana a ese mes completo y "hoy" = su ultimo dia.
  if ($env:FORCE_MES -match '^\d{4}-\d{2}$') {
    $mesIniDt = [DateTime]($env:FORCE_MES + "-01")
    $hoyDt = $mesIniDt.AddMonths(1).AddDays(-1).Date
    Log ("FORZADO mes = " + $env:FORCE_MES + " (hasta " + $hoyDt.ToString("yyyy-MM-dd") + ")")
  }
  $desdeDt = $mesIniDt.AddDays(-21)
  $hoyIso = $hoyDt.ToString("yyyy-MM-dd")

  $cliDias = @{}      # "cliente|fechaEntrega" -> @{fac = ids de boletas; choV; choN}
  $facImp = @{}       # id de boleta -> $ facturado (importeTotal de items, IVA incl.)
  $refImp = @{}       # id de boleta referenciada -> $ rechazado en su contra
  $impRech = @{}      # "mes" -> $ rechazado
  $refMotivo = @{}    # "mes|chofer|idBoletaRef" -> motivo de la primera nota
  $vistosId = @{}
  $zonaSac = @{}; $zonaRech = @{}; $vendSac = @{}; $vendRech = @{}; $cliRechAcum = @{}
  $motivosMes = @{}   # "mes|motivo" -> cantidad de notas
  $repAcum = @{}      # codigoReparto -> lista de {tipo, vd, fp, unidRech} para la efectividad
  $maxFE = ""
  $nVen = 0; $nRech = 0; $nPaginas = 0

  $diaDt = $desdeDt
  while ($diaDt -le $hoyDt) {
    $d1 = $diaDt.ToString("yyyy-MM-dd"); $d2 = $diaDt.AddDays(1).ToString("yyyy-MM-dd")
    $skip = 0
    while ($true) {
      $respPag = Get-Gescom ("ventas/api/v2/get?fechadesde=" + $d1 + "&fechahasta=" + $d2 + "&pagesize=500&pagestoskip=" + $skip + "&pagestotake=1")
      $pagina = @($respPag)
      $nPaginas++
      foreach ($v in $pagina) {
        $tipoV = [string]$v.codigoTipoVenta
        $idV = [string]$v.id
        if (-not $idV -or $vistosId[$idV]) { continue }
        $vistosId[$idV] = $true
        # -- Acumulacion por reparto para la EFECTIVIDAD OFICIAL (todas las ventas
        #    del reparto, aun sin chofer; el chofer lo pone get-repartos despues).
        #    Verificado contra resultado.xlsx: exacto en 220/223 repartos --
        if ($null -ne $v.codigoReparto -and ("" + $v.codigoReparto) -ne "" -and
            ($tipoV -eq "VEN" -or $tipoV -eq "DEV-CA" -or $tipoV -eq "DEV-RE")) {
          $crV = [string]$v.codigoReparto
          if (-not $repAcum[$crV]) { $repAcum[$crV] = New-Object System.Collections.ArrayList }
          $unidRechV = 0.0
          if ($tipoV -eq "DEV-RE") {
            # RechazoItems oficial = unidades x factor de empaque (packs -> unidades)
            foreach ($it in $v.items) {
              $facU = 1.0
              if ($null -ne $it.unidadFactor -and [double]$it.unidadFactor -gt 0) { $facU = [double]$it.unidadFactor }
              $unidRechV += [math]::Abs([double]$it.cantidad) * $facU
            }
          }
          [void]$repAcum[$crV].Add(@{ tipo = $tipoV; vd = ($v.ventaDirecta -eq $true)
                                      fp = ([string]$v.fechaPedido).Substring(0, 10); unidRech = $unidRechV })
        }
        if ($tipoV -ne "VEN" -and $tipoV -ne "DEV-RE") { continue }   # canjes y demas NO cuentan
        $esVenta = ($tipoV -eq "VEN")
        $cho = ""
        if ($v.codigoChofer -and $nomChofer.ContainsKey([string]$v.codigoChofer)) { $cho = $nomChofer[[string]$v.codigoChofer] }
        if (-not $cho -or $cho -in $EXCLUIR) { continue }   # sin chofer asignado = no es reparto de fletero
        if (-not $v.fechaEntrega) { continue }
        $iso = ([string]$v.fechaEntrega).Substring(0, 10)
        if ($iso -gt $hoyIso) { continue }   # entregas futuras precargadas
        $mesK = $iso.Substring(0, 7)
        if ($iso -gt $maxFE) { $maxFE = $iso }
        $codCli = [string]$v.codigoCliente
        $loc = ""; if ($cliLoc.ContainsKey($codCli)) { $loc = $cliLoc[$codCli] }
        $ven = ""; if ($v.codigoVendedor -and $nomVend.ContainsKey([string]$v.codigoVendedor)) { $ven = $nomVend[[string]$v.codigoVendedor] }
        $k = $codCli + "|" + $iso
        if (-not $cliDias[$k]) { $cliDias[$k] = @{ fac = @{}; choV = ""; choN = "" } }
        # Importe por item (con IVA) y reparto por proveedor del articulo
        $impV = 0.0
        foreach ($it in $v.items) {
          $ii = [math]::Abs([double]$it.importeTotal)
          $impV += $ii
          $prov = "Otros"
          $codArt = [string]$it.codigoItem
          if ($provArt.ContainsKey($codArt) -and $provArt[$codArt] -and $nomProv.ContainsKey($provArt[$codArt])) { $prov = $nomProv[$provArt[$codArt]] }
          $kp = "$mesK|$prov"
          $kcp = "$mesK|$cho|$prov"
          if ($esVenta) {
            if (-not $provFact.ContainsKey($kp)) { $provFact[$kp] = 0.0 }
            $provFact[$kp] += $ii
            if (-not $choProvFact.ContainsKey($kcp)) { $choProvFact[$kcp] = 0.0 }
            $choProvFact[$kcp] += $ii
          } else {
            if (-not $provRech.ContainsKey($kp)) { $provRech[$kp] = 0.0 }
            $provRech[$kp] += $ii
            if (-not $choProvRech.ContainsKey($kcp)) { $choProvRech[$kcp] = 0.0 }
            $choProvRech[$kcp] += $ii
          }
        }
        if ($esVenta) {
          $nVen++
          $cliDias[$k].fac[$idV] = $true
          $cliDias[$k].choV = $cho
          $facImp[$idV] = $impV
          if (-not $factTotal.ContainsKey($mesK)) { $factTotal[$mesK] = 0.0 }
          $factTotal[$mesK] += $impV
          if ($loc) { $kz = "$mesK|$loc"; if (-not $zonaSac[$kz]) { $zonaSac[$kz] = @{} }; $zonaSac[$kz][$idV] = $true }
          if ($ven) { $kv = "$mesK|$ven"; if (-not $vendSac[$kv]) { $vendSac[$kv] = @{} }; $vendSac[$kv][$idV] = $true }
          continue
        }
        # --- nota de rechazo (DEV-RE) ---
        $nRech++
        $refId = ""
        if ($v.ventaReferenciada) {
          if ($v.ventaReferenciada.id) { $refId = [string]$v.ventaReferenciada.id }
          elseif ($v.ventaReferenciada.numeroComprobante) { $refId = [string]$v.ventaReferenciada.numeroComprobante }
        }
        if (-not $impRech.ContainsKey($mesK)) { $impRech[$mesK] = 0.0 }
        $impRech[$mesK] += $impV
        if ($refId) {
          if (-not $refImp.ContainsKey($refId)) { $refImp[$refId] = 0.0 }
          $refImp[$refId] += $impV
        }
        $kic = "$mesK|$cho"
        if (-not $impRechCho.ContainsKey($kic)) { $impRechCho[$kic] = 0.0 }
        $impRechCho[$kic] += $impV
        if ($loc -and $refId) { $kz = "$mesK|$loc"; if (-not $zonaRech[$kz]) { $zonaRech[$kz] = @{} }; $zonaRech[$kz][$refId] = $true }
        if ($ven -and $refId) { $kv = "$mesK|$ven"; if (-not $vendRech[$kv]) { $vendRech[$kv] = @{} }; $vendRech[$kv][$refId] = $true }
        $raz = ""; if ($cliRaz.ContainsKey($codCli)) { $raz = $cliRaz[$codCli] }
        if ($raz) {
          $kc = "$mesK|$raz|$loc"
          if (-not $cliRechAcum[$kc]) { $cliRechAcum[$kc] = @{} }
          $cliRechAcum[$kc][$idV] = $true
        }
        $cliDias[$k].choN = $cho
        $mot = ([string]$v.motivo).Trim() -replace "\s+", " "
        if (-not $mot) { $mot = "Sin especificar" }
        $km = "$mesK|$mot"
        if (-not $motivosMes.ContainsKey($km)) { $motivosMes[$km] = 0 }
        $motivosMes[$km]++
        if ($refId) {
          $kr = "$mesK|$cho|$refId"
          if (-not $refMotivo.ContainsKey($kr)) { $refMotivo[$kr] = $mot }
        }
      }
      if ($pagina.Count -lt 500) { break }
      $skip++
      Start-Sleep -Seconds 1
    }
    $diaDt = $diaDt.AddDays(1)
    Start-Sleep -Seconds 1
  }
  Log ("Ventas API OK: " + $nVen + " boletas y " + $nRech + " notas de rechazo en " + $nPaginas +
    " llamadas (cargadas desde " + $desdeDt.ToString("yyyy-MM-dd") + ", entregas hasta " + $maxFE + ")")

  # --- EFECTIVIDAD DE ENTREGA: repartos del mes desde la API ---
  # Reemplaza al resultado*.xlsx (Fase 2, 16/7). Definiciones OFICIALES verificadas
  # reparto por reparto contra el reporte de Gescom (exacto en los 21 fleteros):
  #   Ventas        = boletas VEN (sin venta directa) + canjes DEV-CA
  #                   + devoluciones precargadas ANTES del dia del reparto
  #   RechazoVentas = notas DEV-RE cargadas EN o DESPUES del dia del reparto
  #   RechazoItems  = unidades x factor de empaque de esas notas
  $repartosMes = New-Object System.Collections.ArrayList
  $d1r = $mesIniDt.ToString("yyyy-MM-dd")
  $d2r = $hoyDt.AddDays(1).ToString("yyyy-MM-dd")   # fechahasta exclusiva -> incluye hoy
  $skipR = 0
  while ($true) {
    $respR = Get-Gescom ("distribucion/api/v1/get-repartos?fechadesde=" + $d1r + "&fechahasta=" + $d2r + "&pagesize=500&pagestoskip=" + $skipR + "&pagestotake=1")
    $pagR = @($respR)
    foreach ($rp in $pagR) { [void]$repartosMes.Add($rp) }
    if ($pagR.Count -lt 500) { break }
    $skipR++
    Start-Sleep -Seconds 1
  }
  $entregas = @{}   # clave "fecha|CHOFER" -> @{asig; real; itemsRech}
  $repartosCho = @{} # CHOFER -> cantidad de repartos hechos en el mes (tarjeta del detalle)
  foreach ($rp in $repartosMes) {
    $crR = [string]$rp.codigo
    if (-not $rp.fecha) { continue }
    $fechaR = ([string]$rp.fecha).Substring(0, 10)
    if ($fechaR -gt $hoyIso) { continue }   # repartos futuros precargados
    $choR = ""
    if ($rp.codigoChofer -and $nomChofer.ContainsKey([string]$rp.codigoChofer)) { $choR = $nomChofer[[string]$rp.codigoChofer] }
    if (-not $choR -and $rp.nombreChofer) { $choR = (([string]$rp.nombreChofer).Trim().ToUpper() -replace "\s+", " ") }
    if (-not $choR -or $choR -in $EXCLUIR) { continue }
    $lista = $repAcum[$crR]
    if (-not $lista) { continue }
    $asigR = 0; $rechR = 0; $itemsR = 0.0
    foreach ($mv in $lista) {
      if ($mv.tipo -eq "DEV-RE") {
        if ($mv.fp -lt $fechaR) { $asigR++ }                    # devolucion precargada: cuenta como venta
        else { $rechR++; $itemsR += $mv.unidRech }              # nota del reparto: rechazo
      } elseif ($mv.tipo -eq "VEN") {
        if (-not $mv.vd) { $asigR++ }                           # las "ventas directas" no cuentan
      } else {
        $asigR++                                                # canje DEV-CA
      }
    }
    if ($asigR -le 0) { continue }
    if (-not $repartosCho[$choR]) { $repartosCho[$choR] = 0 }
    $repartosCho[$choR]++
    $claveR = "$fechaR|$choR"
    if (-not $entregas[$claveR]) { $entregas[$claveR] = @{ asig = 0; real = 0; itemsRech = 0; rep = 0 } }
    $entregas[$claveR].rep += 1
    $entregas[$claveR].asig += $asigR
    $entregas[$claveR].real += [math]::Max(0, $asigR - $rechR)
    $entregas[$claveR].itemsRech += [int][math]::Round($itemsR)
  }
  $choferesGescom = @($entregas.Keys | ForEach-Object { $_.Split("|")[1] } | Sort-Object -Unique)
  Log ("Efectividad API OK: " + $repartosMes.Count + " repartos del mes -> " + $entregas.Count +
    " registros dia/chofer, " + $choferesGescom.Count + " choferes")

  # --- Mes en curso = mes de la ultima entrega ---
  $mesFE = ""
  if ($maxFE) { $mesFE = $maxFE.Substring(0, 7) }

  # Motivos del mes (tarjeta general: cuenta TODAS las notas)
  foreach ($km in @($motivosMes.Keys)) {
    $pm = $km.Split("|", 2)
    if ($pm[0] -ne $mesFE) { continue }
    $motivos[$pm[1]] = $motivosMes[$km]
  }

  # --- Estadisticas del mes por chofer: rechazos totales/parciales, clientes y boletas ---
  foreach ($k in @($cliDias.Keys)) {
    if ($mesFE -and $k.Split("|")[1] -notlike "$mesFE*") { continue }   # solo mes en curso
    $d = $cliDias[$k]
    $cho = $d.choV; if (-not $cho) { $cho = $d.choN }
    if (-not $cho) { continue }
    if (-not $statsChofer[$cho]) {
      $statsChofer[$cho] = @{ recTot = 0; recBol = 0; cliSac = 0; compSac = 0; compRech = 0; prodSuel = 0 }
    }
    $s = $statsChofer[$cho]
    $boletas = $d.fac.Count
    if ($boletas -gt 0) {
      # Boleta rechazada COMPLETA: las notas cubren (casi) todo el importe de la boleta.
      # Un rechazo de productos sueltos NO cuenta contra el fletero.
      $bolComp = 0
      foreach ($fx in @($d.fac.Keys)) {
        $fi = 0.0; if ($facImp.ContainsKey($fx)) { $fi = $facImp[$fx] }
        $ri = 0.0; if ($refImp.ContainsKey($fx)) { $ri = $refImp[$fx] }
        if ($fi -gt 0 -and $ri -ge (0.98 * $fi)) { $bolComp++ }
      }
      $s.cliSac++
      $s.compSac += $boletas
      $s.compRech += $bolComp
      $s.recBol += $bolComp
      if ($bolComp -ge $boletas) { $s.recTot++ }   # cliente completo: TODAS sus boletas enteras
    }
  }

  # Motivos por fletero: SOLO de boletas rechazadas completas (una entrada por boleta,
  # asi el total de la tabla coincide con "boletas rechazadas completas")
  $motivosPorChofer = @{}
  foreach ($kr in @($refMotivo.Keys)) {
    $pp = $kr.Split("|")
    if ($pp[0] -ne $mesFE) { continue }
    $cho2 = $pp[1]; $ref2 = $pp[2]
    $fi = 0.0; if ($facImp.ContainsKey($ref2)) { $fi = $facImp[$ref2] }
    $ri = 0.0; if ($refImp.ContainsKey($ref2)) { $ri = $refImp[$ref2] }
    if (-not ($fi -gt 0 -and $ri -ge (0.98 * $fi))) { continue }   # solo boletas completas
    $mot2 = $refMotivo[$kr]
    if (-not $mot2) { $mot2 = "Sin especificar" }
    if (-not $motivosPorChofer[$cho2]) { $motivosPorChofer[$cho2] = @{} }
    if (-not $motivosPorChofer[$cho2][$mot2]) { $motivosPorChofer[$cho2][$mot2] = 0 }
    $motivosPorChofer[$cho2][$mot2]++
  }
  Log ("Motivos OK: " + ($motivos.Values | Measure-Object -Sum).Sum + " rechazos (notas), " +
    $motivos.Count + " motivos, " + $motivosPorChofer.Count + " choferes con detalle")
  Log ("Estadisticas de mes (" + $mesFE + "): " + $statsChofer.Count + " choferes con clientes/boletas/rechazos tot-par")

  # --- Analisis de rechazos del mes: zonas, vendedores, clientes, importe ---
  function Top-Porcentaje($sacMap, $rechMap, $mes, $minBoletas) {
    $lista = foreach ($kk in @($sacMap.Keys)) {
      if ($kk -notlike "$mes|*") { continue }
      $sac = $sacMap[$kk].Count
      if ($sac -lt $minBoletas) { continue }
      # Misma vara que el fletero: solo cuentan boletas rechazadas COMPLETAS
      # (venta caida de verdad; un producto suelto no voltea la venta)
      $rech = 0
      if ($rechMap[$kk]) {
        foreach ($rf in @($rechMap[$kk].Keys)) {
          $fi = 0.0; if ($facImp.ContainsKey($rf)) { $fi = $facImp[$rf] }
          $ri = 0.0; if ($refImp.ContainsKey($rf)) { $ri = $refImp[$rf] }
          if ($fi -gt 0 -and $ri -ge (0.98 * $fi)) { $rech++ }
        }
        $rech = [math]::Min($rech, $sac)
      }
      if ($rech -eq 0) { continue }
      [PSCustomObject]@{ nombre = $kk.Split("|")[1]; sac = $sac; rech = $rech; pct = [math]::Round(100.0 * $rech / $sac, 1) }
    }
    return @($lista | Sort-Object pct -Descending | Select-Object -First 8)
  }
  $anZonas = Top-Porcentaje $zonaSac $zonaRech $mesFE 20
  $anVend = Top-Porcentaje $vendSac $vendRech $mesFE 20
  $anClientes = @(foreach ($kk in @($cliRechAcum.Keys)) {
    if ($kk -notlike "$mesFE|*") { continue }
    $p = $kk.Split("|")
    [PSCustomObject]@{ nombre = $p[1]; loc = $p[2]; cantidad = $cliRechAcum[$kk].Count }
  }) | Sort-Object cantidad -Descending | Select-Object -First 8
  $anImporte = 0
  if ($impRech.ContainsKey($mesFE)) { $anImporte = [math]::Round($impRech[$mesFE]) }
  if ($factTotal.ContainsKey($mesFE)) { $anFacturado = [math]::Round($factTotal[$mesFE]) }
  # % entregado en plata por proveedor (empresa; min $1M facturado)
  $anProveedores = @(foreach ($kp in @($provFact.Keys)) {
    if ($kp -notlike "$mesFE|*") { continue }
    $fv = $provFact[$kp]
    if ($fv -lt 1000000) { continue }
    $rv = 0.0; if ($provRech.ContainsKey($kp)) { $rv = $provRech[$kp] }
    [PSCustomObject]@{ nombre = $kp.Split("|")[1]; fac = [math]::Round($fv); rech = [math]::Round($rv); pct = [math]::Round(100.0 * ($fv - $rv) / $fv, 1) }
  }) | Sort-Object fac -Descending | Select-Object -First 8
  Log ("Analisis rechazos: " + $anZonas.Count + " zonas, " + $anVend.Count + " vendedores, " +
    $anClientes.Count + " clientes top, importe total `$" + $anImporte)
} catch {
  Log ("ERROR leyendo la API de Gescom: " + $_.Exception.Message)
  Log "No se publica nada: el panel queda como estaba. Probar de nuevo mas tarde con el acceso directo."
  exit 1
}

# ============================================================================
# 2) CARTON -> retorno de carton
# ============================================================================
# La planilla se baja directo de Google Drive (link "cualquiera con el enlace
# puede ver", guardado como secreto CARTON_URL). En modo prueba local, usa el
# ultimo Carton*.xlsx descargado a Documentos\GESCOM.
$rutaCarton = ""
if ($env:CARTON_URL) {
  $rutaCarton = Join-Path $env:TEMP "carton-drive.xlsx"
  try {
    Invoke-WebRequest -Uri $env:CARTON_URL -OutFile $rutaCarton -TimeoutSec 120 -UseBasicParsing
  } catch {
    Log ("ERROR bajando la planilla de carton de Drive: " + $_.Exception.Message)
    Log "No se publica nada: el panel queda como estaba."
    exit 1
  }
  $tamanoKb = [math]::Round((Get-Item $rutaCarton).Length / 1024)
  if ($tamanoKb -lt 5) { Log "ERROR: la descarga del carton vino vacia (revisar el link CARTON_URL)"; exit 1 }
  Log ("Carton: descargado de Drive (" + $tamanoKb + " KB)")
} else {
  $archCarton = Get-ChildItem "C:\Users\luqaa\Documents\GESCOM\Carton*.xlsx" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $archCarton) { Log "ERROR: sin CARTON_URL y no encuentro Carton*.xlsx local"; exit 1 }
  $rutaCarton = $archCarton.FullName
  Log ("Carton (prueba local): " + $archCarton.Name + " (modificado " + $archCarton.LastWriteTime.ToString("dd/MM HH:mm") + ")")
}

# Mapeo nombre corto ("Carlos C") -> nombre completo Gescom ("CARLOS CRESPO")
$overrides = @{}
$archMapeo = Join-Path $RAIZ "mapeo-nombres.txt"
if (Test-Path $archMapeo) {
  foreach ($l in Get-Content $archMapeo -Encoding UTF8) {
    if ($l -match "^\s*([^=#]+?)\s*=\s*(.+?)\s*$") { $overrides[$Matches[1].ToUpper()] = $Matches[2].ToUpper() }
  }
}
$sinMapear = New-Object System.Collections.ArrayList
function Mapear-Nombre($corto) {
  $c = $corto.Trim()
  $cU = $c.ToUpper()
  if ($overrides[$cU]) { return $overrides[$cU] }
  $partes = $cU -split "\s+"
  if ($partes.Count -ge 2) {
    $nombre = $partes[0]; $inicial = $partes[1].Substring(0, 1)
    $candidatos = @($choferesGescom | Where-Object {
      $p = $_ -split "\s+"
      $p[0] -eq $nombre -and $p.Count -ge 2 -and $p[1].StartsWith($inicial)
    })
    if ($candidatos.Count -eq 1) { return $candidatos[0] }
  }
  if ($sinMapear -notcontains $c) { [void]$sinMapear.Add($c) }
  return $cU
}

$MESES = @{ "Enero"=1; "Febrero"=2; "Marzo"=3; "Abril"=4; "Mayo"=5; "Junio"=6; "Julio"=7; "Agosto"=8; "Septiembre"=9; "Octubre"=10; "Noviembre"=11; "Diciembre"=12 }
$libroC = Abrir-Xlsx $rutaCarton
$cartones = @{}   # clave "fecha|CHOFER" -> @{sal; vue}
$vistos = @{}     # dedupe exacto fecha|fletero|reparto
$fechaMin = (Get-Date).AddDays(-$DIAS_HISTORIAL)

foreach ($hoja in @($libroC.hojas.Keys)) {
  if (-not $MESES.ContainsKey($hoja)) { continue }   # salta ModeloEnBlanco, Semanal Mayo, etc.
  $filas = Leer-Hoja $libroC $hoja
  $mapa = $null
  $n = 0
  foreach ($fila in $filas) {
    # Encabezado de bloque diario: "Fecha" en columna A (la tabla celeste tiene el chofer en A, no pasa este filtro)
    if (([string]$fila["A"]).Trim() -eq "Fecha") {
      $mapa = @{}
      foreach ($k in @($fila.Keys)) {
        switch -Regex (([string]$fila[$k]).Trim()) {
          "^Fletero"    { $mapa.fletero = $k }
          "^Reparto$"   { $mapa.reparto = $k }
          "^Salida$"    { $mapa.salida = $k }
          "^Vuelve$"    { $mapa.vuelve = $k }
        }
      }
      continue
    }
    if (-not $mapa -or -not (EsNumero $fila["A"])) { continue }   # solo filas con fecha-numero de Excel
    $serial = [double]$fila["A"]
    if ($serial -lt 40000 -or $serial -gt 60000) { continue }
    $fechaDt = [DateTime]::FromOADate($serial)
    if ($fechaDt -lt $fechaMin) { continue }
    $fletero = ([string]$fila[$mapa.fletero]).Trim()
    if (-not $fletero) { continue }
    if (-not (EsNumero $fila[$mapa.salida]) -or -not (EsNumero $fila[$mapa.vuelve])) { continue }
    $sal = [int][double]$fila[$mapa.salida]
    $vue = [int][double]$fila[$mapa.vuelve]
    if ($sal -le 0) { continue }   # SinReparto / PeñaFlor / Palett = no salio a repartir
    $reparto = [string]$fila[$mapa.reparto]
    $claveDedupe = $fechaDt.ToString("yyyy-MM-dd") + "|" + $fletero.ToUpper() + "|" + $reparto
    if ($vistos[$claveDedupe]) { continue }
    $vistos[$claveDedupe] = $true
    $completo = Mapear-Nombre $fletero
    $clave = $fechaDt.ToString("yyyy-MM-dd") + "|" + $completo
    if (-not $cartones[$clave]) { $cartones[$clave] = @{ sal = 0; vue = 0 } }
    $cartones[$clave].sal += $sal
    $cartones[$clave].vue += [math]::Min($vue, $sal)   # tope 100%: no puede volver mas de lo que salio
    $n++
  }
  if ($n -gt 0) { Log ("Carton hoja '$hoja': $n filas validas") }
}
Cerrar-Xlsx $libroC
if ($sinMapear.Count -gt 0) { Log ("AVISO nombres sin mapear (agregalos a mapeo-nombres.txt): " + ($sinMapear -join ", ")) }
Log ("Carton OK: " + $cartones.Count + " registros dia/fletero")

# ============================================================================
# 3) Unir y generar data.js
# ============================================================================
$claves = @($entregas.Keys) + @($cartones.Keys) | Sort-Object -Unique
# Ignorar fechas futuras (Gescom trae repartos ya cargados para dias que no pasaron)
$hoy = (Get-Date).ToString("yyyy-MM-dd")
$claves = @($claves | Where-Object { $_.Split("|")[0] -le $hoy })
# MODO VERIFICACION: nos quedamos con TODO el mes forzado y nada mas.
# Sin esto pasan dos desastres (ya ocurrieron el 20/8):
#   a) el recorte de "ultimas N fechas" de mas abajo se come el principio del
#      mes -> la asistencia se desploma y NADIE cobra premio;
#   b) el carton de la planilla trae tambien el mes siguiente y ensucia el total.
if ($env:FORCE_MES) {
  $claves = @($claves | Where-Object { $_.Split("|")[0] -like ($env:FORCE_MES + "*") })
  if ($claves.Count -eq 0) { Log ("ERROR: no hay datos del mes " + $env:FORCE_MES); exit 1 }
}
$fechasTodas = @($claves | ForEach-Object { $_.Split("|")[0] } | Sort-Object -Unique)
# Publicar las ultimas N fechas con datos: minimo 14 (para los promedios 14d)
# y ademas todo el mes en curso (para los totales mensuales por fletero)
$ultimaFecha = $fechasTodas[-1]
$diaDelMes = [int]$ultimaFecha.Substring(8, 2)
$nPublicar = [Math]::Max($DIAS_PUBLICAR, $diaDelMes)
if ($env:FORCE_MES) { $nPublicar = $fechasTodas.Count }   # verificacion: el mes ENTERO
$fechasPublicar = @($fechasTodas | Select-Object -Last $nPublicar)
$claves = @($claves | Where-Object { $_.Split("|")[0] -in $fechasPublicar })
Log ("Datos publicados: " + $fechasPublicar[0] + " a " + $fechasPublicar[-1] + " (" + $fechasPublicar.Count + " fechas)")

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("/* GENERADO AUTOMATICAMENTE por robot-actualizar-web.ps1 - NO EDITAR A MANO")
[void]$sb.AppendLine("   Ultima actualizacion: " + (Get-Date -Format "yyyy-MM-dd HH:mm") + " */")
[void]$sb.AppendLine("window.__PPP_CONFIG__ = {")
[void]$sb.AppendLine('  SHEET_CSV_URL: "",')
[void]$sb.AppendLine("  umbrales: { bueno: 90, medio: 75 },")
[void]$sb.AppendLine("  diasHistorial: 14")
[void]$sb.AppendLine("};")
[void]$sb.AppendLine("window.__PPP_DATA__ = { registros: [")
$primero = $true
foreach ($clave in $claves) {
  $p = $clave.Split("|"); $fecha = $p[0]; $chofer = $p[1]
  if ($chofer -in $EXCLUIR) { continue }
  $e = $entregas[$clave]; $c = $cartones[$clave]
  $ea = 0; $er = 0; $ca = 0; $cr = 0; $nrep = 0
  if ($e) { $ea = $e.asig; $er = $e.real; if ($e.rep) { $nrep = $e.rep } }
  if ($c) { $ca = $c.sal; $cr = $c.vue }
  # Nombre para mostrar: "Carlos Crespo" en vez de "CARLOS CRESPO"
  $mostrar = (($chofer.ToLower() -split "\s+") | ForEach-Object { if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) } }) -join " "
  $coma = ","; if ($primero) { $coma = " "; $primero = $false }
  $json = '{"fecha":"' + $fecha + '","fletero":"' + $mostrar + '","zona":"","repartos":' + $nrep + ',"entregas_asignadas":' + $ea + ',"entregas_realizadas":' + $er + ',"cartones_a_retornar":' + $ca + ',"cartones_retornados":' + $cr + '}'
  [void]$sb.AppendLine($coma + $json)
}
[void]$sb.AppendLine("] };")
# Motivos de rechazo (ordenados de mas a menos frecuente)
$listaMot = $motivos.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
  '{"motivo":"' + ($_.Key -replace '"', "'") + '","cantidad":' + $_.Value + '}'
}
[void]$sb.AppendLine("window.__PPP_DATA__.motivos = [" + ($listaMot -join ",") + "];")
# Feriados del anio (para descontarlos de los dias habiles en la asistencia)
$listaFer = @($feriadosWeb | ForEach-Object { '"' + $_ + '"' }) -join ","
[void]$sb.AppendLine("window.__PPP_DATA__.feriados = [" + $listaFer + "];")
# Motivos por fletero (clave = nombre para mostrar, igual que en registros)
function NombreMostrar($chofer) {
  (($chofer.ToLower() -split "\s+") | ForEach-Object { if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) } }) -join " "
}
$porFle = foreach ($cho in ($motivosPorChofer.Keys | Sort-Object)) {
  $lista = $motivosPorChofer[$cho].GetEnumerator() | Sort-Object Value -Descending | ForEach-Object {
    '{"motivo":"' + ($_.Key -replace '"', "'") + '","cantidad":' + $_.Value + '}'
  }
  '"' + (NombreMostrar $cho) + '":[' + ($lista -join ",") + ']'
}
[void]$sb.AppendLine("window.__PPP_DATA__.motivosPorFletero = {" + ($porFle -join ",") + "};")
# Estadisticas del mes por fletero (rechazos totales/parciales, clientes y boletas)
# Items (productos) rechazados del mes por chofer, desde el reporte oficial de Gescom
$mesStats = $ultimaFecha.Substring(0, 7)
$itemsRechMes = @{}
foreach ($claveI in $claves) {
  if ($claveI.Split("|")[0] -notlike "$mesStats*") { continue }
  $choI = $claveI.Split("|")[1]
  $eI = $entregas[$claveI]
  if ($eI -and $eI.itemsRech) {
    if (-not $itemsRechMes[$choI]) { $itemsRechMes[$choI] = 0 }
    $itemsRechMes[$choI] += $eI.itemsRech
  }
}
$statsJson = foreach ($cho in ($statsChofer.Keys | Sort-Object)) {
  $s = $statsChofer[$cho]
  $pSuel = 0; if ($s.prodSuel) { $pSuel = $s.prodSuel }
  $iRech = 0; if ($itemsRechMes[$cho]) { $iRech = $itemsRechMes[$cho] }
  $iImpR = 0; if ($impRechCho.ContainsKey("$mesFE|$cho")) { $iImpR = [math]::Round($impRechCho["$mesFE|$cho"]) }
  $nRep = 0; if ($repartosCho[$cho]) { $nRep = $repartosCho[$cho] }
  '"' + (NombreMostrar $cho) + '":{"recTot":' + $s.recTot + ',"recBol":' + $s.recBol +
    ',"prodSuel":' + $pSuel + ',"itemsRech":' + $iRech + ',"impRech":' + $iImpR +
    ',"repartos":' + $nRep +
    ',"cliSac":' + $s.cliSac + ',"cliEnt":' + ($s.cliSac - $s.recTot) +
    ',"compSac":' + $s.compSac + ',"compEnt":' + ($s.compSac - $s.compRech) + '}'
}
[void]$sb.AppendLine("window.__PPP_DATA__.estadisticasFletero = {" + ($statsJson -join ",") + "};")
# Analisis de rechazos del mes (zonas, vendedores, clientes, importe)
function JsonTxt($s) { return ([string]$s -replace '\\', '\\\\' -replace '"', "'") }
$jZonas = @($anZonas | ForEach-Object { '{"nombre":"' + (JsonTxt $_.nombre) + '","pct":' + ($_.pct -replace ",", ".") + ',"rech":' + $_.rech + ',"sac":' + $_.sac + '}' }) -join ","
$jVend = @($anVend | ForEach-Object { '{"nombre":"' + (JsonTxt $_.nombre) + '","pct":' + ($_.pct -replace ",", ".") + ',"rech":' + $_.rech + ',"sac":' + $_.sac + '}' }) -join ","
$jCli = @($anClientes | ForEach-Object { '{"nombre":"' + (JsonTxt $_.nombre) + '","loc":"' + (JsonTxt $_.loc) + '","cantidad":' + $_.cantidad + '}' }) -join ","
# Privacidad: NINGUN monto se publica (ni facturado ni rechazado); solo porcentajes y cantidades
$jProv = @($anProveedores | ForEach-Object { '{"nombre":"' + (JsonTxt $_.nombre) + '","pct":' + ($_.pct -replace ",", ".") + '}' }) -join ","
[void]$sb.AppendLine('window.__PPP_DATA__.analisisRechazos = {"importe":' + $anImporte + ',"zonas":[' + $jZonas + '],"vendedores":[' + $jVend + '],"clientes":[' + $jCli + '],"proveedores":[' + $jProv + ']};')
# Entrega por proveedor de cada fletero (en plata, min $100k, top 6)
$porCho3 = @{}
foreach ($kcp in @($choProvFact.Keys)) {
  $pp3 = $kcp.Split("|")
  if ($pp3[0] -ne $mesFE) { continue }
  $cho3 = $pp3[1]; $pr3 = $pp3[2]
  if ($cho3 -in $EXCLUIR) { continue }
  $fv = $choProvFact[$kcp]
  if ($fv -lt 100000) { continue }
  $rv = 0.0; if ($choProvRech.ContainsKey($kcp)) { $rv = $choProvRech[$kcp] }
  if (-not $porCho3[$cho3]) { $porCho3[$cho3] = New-Object System.Collections.ArrayList }
  [void]$porCho3[$cho3].Add([PSCustomObject]@{ prov = $pr3; fac = [math]::Round($fv); pct = [math]::Round(100.0 * ($fv - $rv) / $fv, 1) })
}
$jFleProv = foreach ($cho3 in ($porCho3.Keys | Sort-Object)) {
  $lst = @($porCho3[$cho3] | Sort-Object fac -Descending | Select-Object -First 6 | ForEach-Object {
    '{"prov":"' + (JsonTxt $_.prov) + '","pct":' + ($_.pct -replace ",", ".") + '}'
  })
  '"' + (NombreMostrar $cho3) + '":[' + ($lst -join ",") + ']'
}
[void]$sb.AppendLine("window.__PPP_DATA__.proveedoresPorFletero = {" + ($jFleProv -join ",") + "};")

# --- Tarjeta de cierre de mes: emitir el MES ANTERIOR (ya cerrado) ----------
# Sale del historial-meses.json (esquema rico). Se lee ACA porque va dentro de
# data.js; el historial del mes en curso se (re)escribe mas abajo (seccion 3b).
$mesActual = $mesFE; if (-not $mesActual) { $mesActual = (Get-Date -Format "yyyy-MM") }
$histFile = Join-Path $RAIZ "historial-meses.json"
$hist = @{}
if (Test-Path $histFile) {
  try {
    $viejo = Get-Content $histFile -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($p in $viejo.PSObject.Properties) { $hist[$p.Name] = $p.Value }
  } catch { Log "AVISO: no pude leer historial-meses.json, se regenera" }
}
function NumH($v) { if ($null -eq $v) { return 0 } else { return $v } }
# En modo verificacion (FORCE_MES) guardamos el registro VIEJO del mes forzado
# ANTES de recalcularlo, para comparar despues (seccion 3b) y ver si aparecieron
# modificaciones tardias (notas cargadas despues del cierre).
$verifMesViejo = $null
if ($env:FORCE_MES -match '^\d{4}-\d{2}$') { $verifMesViejo = $hist[$env:FORCE_MES] }
# El mes anterior es el mes CALENDARIO anterior a hoy (no el del ultimo dato):
# asi la tarjeta celebra el mes recien cerrado aunque el mes nuevo todavia no
# tenga entregas cargadas. Con FORCE_MES no importa (no usamos su data.js).
$mesAntKey = (Get-Date -Day 1).Date.AddMonths(-1).ToString("yyyy-MM")
$maH = $hist[$mesAntKey]
if ($maH -and $null -ne $maH.ranking) {
  $maRank = @()
  foreach ($rk in @($maH.ranking)) {
    if (([string]$rk.nombre).ToUpper() -in $EXCLUIR) { continue }   # ocultar excluidos
    $efEj = "null"; if ($null -ne $rk.efE) { $efEj = ([string]$rk.efE) -replace ",", "." }
    $efCj = "null"; if ($null -ne $rk.efC) { $efCj = ([string]$rk.efC) -replace ",", "." }
    # asist puede faltar en meses guardados antes (la web la recalcula sola)
    $asJ = "null"; if ($null -ne $rk.asist) { $asJ = ([string]$rk.asist) -replace ",", "." }
    $maRank += '{"nombre":"' + (JsonTxt $rk.nombre) + '","repartos":' + ([int](NumH $rk.repartos)) +
      ',"asist":' + $asJ + ',"efE":' + $efEj + ',"efC":' + $efCj + ',"premio":' + ([long](NumH $rk.premio)) + '}'
  }
  $maJson = '{"clave":"' + $mesAntKey + '","anio":' + [int]$mesAntKey.Substring(0, 4) + ',"mes":' + [int]$mesAntKey.Substring(5, 2) +
    ',"efGeneral":' + ((([string](NumH $maH.efGeneral))) -replace ",", ".") +
    ',"cartonGeneral":' + ((([string](NumH $maH.cartonGeneral))) -replace ",", ".") +
    ',"repartos":' + [int](NumH $maH.repartos) +
    ',"boletasEnt":' + [int](NumH $maH.boletasEnt) + ',"boletasSac":' + [int](NumH $maH.boletasSac) +
    ',"clientesEnt":' + [int](NumH $maH.clientesEnt) + ',"clientesSac":' + [int](NumH $maH.clientesSac) +
    ',"plataRech":' + [long](NumH $maH.plataRech) + ',"fleteros":' + [int](NumH $maH.fleteros) +
    ',"premiosTotal":' + [long](NumH $maH.premiosTotal) +
    ',"ranking":[' + ($maRank -join ",") + ']}'
  [void]$sb.AppendLine("window.__PPP_DATA__.mesAnterior = " + $maJson + ";")
  Log ("Tarjeta de cierre: mes anterior " + $mesAntKey + " emitido (" + $maRank.Count + " fleteros)")
} else {
  Log ("Tarjeta de cierre: sin datos ricos del mes anterior (" + $mesAntKey + ") todavia")
}

$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $RAIZ "data.js"), $sb.ToString(), $utf8)
Log "data.js generado"

# ============================================================================
# 3b) Historial mensual RICO (alimenta la tarjeta de cierre del mes siguiente)
#     Guarda por mes: efectividad y carton generales, repartos, boletas y
#     clientes, plata rechazada, premios totales, y el ranking por fletero
#     (con entrega, carton y premio de cada uno). $hist y $histFile ya se
#     leyeron arriba (seccion de la tarjeta de cierre).
# ============================================================================
# Dias habiles (lun-vie) del mes cerrado; los feriados NO se descuentan (regla
# de la empresa: el feriado no trabajado se compensa repartiendo el sabado).
$habMes = 0
$dHab = [DateTime]($mesActual + "-01")
while ($dHab.ToString("yyyy-MM") -eq $mesActual) {
  $dow = [int]$dHab.DayOfWeek
  if ($dow -ge 1 -and $dow -le 5) { $habMes++ }
  $dHab = $dHab.AddDays(1)
}
# Agregacion del mes por fletero (excluye no-fleteros y los que Lucas oculta)
$flAcum = @{}
foreach ($clave in $claves) {
  if ($clave.Split("|")[0] -notlike "$mesActual*") { continue }
  $choH = $clave.Split("|")[1]
  if ($choH -in $EXCLUIR) { continue }
  if (-not $flAcum[$choH]) { $flAcum[$choH] = @{ asig = 0; real = 0; sal = 0; vue = 0; rep = 0 } }
  $eH = $entregas[$clave]; $cH = $cartones[$clave]
  if ($eH) { $flAcum[$choH].asig += $eH.asig; $flAcum[$choH].real += $eH.real; if ($eH.rep) { $flAcum[$choH].rep += $eH.rep } }
  if ($cH) { $flAcum[$choH].sal += $cH.sal; $flAcum[$choH].vue += $cH.vue }
}
$mEA = 0; $mER = 0; $mCA = 0; $mCR = 0; $mRep = 0; $premTot = 0
$rankArr = @()
foreach ($choH in ($flAcum.Keys | Sort-Object)) {
  $a = $flAcum[$choH]
  # El carton siempre suma al total de la empresa (aunque el nombre de la
  # planilla no haya mapeado a un chofer); asi coincide con el anillo del panel.
  $mCA += $a.sal; $mCR += $a.vue
  # Pero el ranking/premios son SOLO de fleteros con repartos reales: una fila
  # de carton sin repartos (nombre sin mapear) no es un fletero rankeable.
  if ($a.rep -le 0) { continue }
  $mEA += $a.asig; $mER += $a.real; $mRep += $a.rep
  $efEh = $null; if ($a.asig -gt 0) { $efEh = PctEntero $a.real $a.asig }
  $efCh = $null; if ($a.sal -gt 0) { $efCh = PctEntero $a.vue $a.sal }
  # Asistencia = repartos / dias habiles (tope 100%); <85% no cobra premio
  $asistH = $null; if ($habMes -gt 0) { $asistH = [math]::Min(100, (PctEntero $a.rep $habMes)) }
  $premH = 0
  if ($null -ne $asistH -and $asistH -ge 85) {
    if ($null -ne $efEh) { if ($efEh -ge 95) { $premH += 100000 } elseif ($efEh -ge 90) { $premH += 50000 } }
    if ($null -ne $efCh) { if ($efCh -ge 80) { $premH += 150000 } elseif ($efCh -ge 70) { $premH += 100000 } elseif ($efCh -ge 60) { $premH += 50000 } }
  }
  $premTot += $premH
  $rankArr += [PSCustomObject]@{ nombre = (NombreMostrar $choH); repartos = $a.rep; asist = $asistH; efE = $efEh; efC = $efCh; premio = $premH }
}
$rankArr = @($rankArr | Sort-Object { if ($null -eq $_.efE) { -1.0 } else { [double]$_.efE } } -Descending)
$efGenH = 0; if ($mEA -gt 0) { $efGenH = PctEntero $mER $mEA }
$carGenH = 0; if ($mCA -gt 0) { $carGenH = PctEntero $mCR $mCA }
# Clientes del mes (de statsChofer, que ya excluye no-fleteros)
$cliSacT = 0; $cliEntT = 0
foreach ($cho in $statsChofer.Keys) {
  $s = $statsChofer[$cho]
  $cliSacT += $s.cliSac; $cliEntT += ($s.cliSac - $s.recTot)
}
$hist[$mesActual] = @{
  efGeneral = $efGenH; cartonGeneral = $carGenH
  boletasSac = $mEA; boletasEnt = $mER
  clientesSac = $cliSacT; clientesEnt = $cliEntT
  repartos = $mRep; plataRech = $anImporte
  premiosTotal = $premTot; fleteros = $rankArr.Count
  ranking = $rankArr
  actualizado = (Get-Date -Format "yyyy-MM-dd")
}
[System.IO.File]::WriteAllText($histFile, ($hist | ConvertTo-Json -Depth 6), $utf8)
Log ("Historial mensual actualizado (" + $mesActual + "; " + $flAcum.Count + " fleteros, premios `$" + $premTot + "; meses: " + $hist.Count + ")")

# ============================================================================
# 3c) MODO VERIFICACION (FORCE_MES): comparar el mes recalculado contra su
#     version anterior y escribir un informe en criollo (verificacion-mes.md).
#     Sirve para corroborar que ya entraron todas las modificaciones tardias.
# ============================================================================
if ($env:FORCE_MES -match '^\d{4}-\d{2}$') {
  $vn = $hist[$mesActual]   # nuevo (recien calculado)
  $vv = $verifMesViejo      # viejo (antes de recalcular; puede ser $null)
  function Fnum($x) { if ($null -eq $x) { return 0 } else { return [long]$x } }
  $lineas = New-Object System.Collections.ArrayList
  [void]$lineas.Add("# Verificacion del cierre de " + $mesActual)
  [void]$lineas.Add("")
  [void]$lineas.Add("Recalculado el " + (Get-Date -Format "yyyy-MM-dd HH:mm") + " (hora argentina), anclado en la fecha de reparto (cierre de camiones).")
  [void]$lineas.Add("")
  if (-not $vv) {
    [void]$lineas.Add("No habia una version anterior guardada para comparar; se dejaron los numeros actuales como definitivos.")
  } else {
    $cambios = New-Object System.Collections.ArrayList
    $campos = @(
      @("Efectividad de entrega", "efGeneral", "%"),
      @("Retorno de carton", "cartonGeneral", "%"),
      @("Repartos", "repartos", ""),
      @("Boletas entregadas", "boletasEnt", ""),
      @("Boletas asignadas", "boletasSac", ""),
      @("Clientes entregados", "clientesEnt", ""),
      @("Clientes visitados", "clientesSac", ""),
      @("Plata rechazada", "plataRech", "`$"),
      @("Premios totales", "premiosTotal", "`$"),
      @("Fleteros", "fleteros", "")
    )
    foreach ($cp in $campos) {
      $antes = $vv.($cp[1]); $desp = $vn[$cp[1]]
      if (([string]$antes) -ne ([string]$desp)) {
        [void]$cambios.Add("- **" + $cp[0] + "**: " + $cp[2] + [string]$antes + " -> " + $cp[2] + [string]$desp)
      }
    }
    # Premios por fletero que cambiaron
    $premViejo = @{}
    if ($null -ne $vv.ranking) { foreach ($rk in @($vv.ranking)) { $premViejo[[string]$rk.nombre] = [long](Fnum $rk.premio) } }
    $premCambio = New-Object System.Collections.ArrayList
    foreach ($rk in $vn.ranking) {
      $nom = [string]$rk.nombre
      $pv = 0; if ($premViejo.ContainsKey($nom)) { $pv = $premViejo[$nom] }
      if ($pv -ne [long]$rk.premio) {
        [void]$premCambio.Add("- **" + $nom + "**: premio `$" + $pv + " -> `$" + [long]$rk.premio)
      }
    }
    if ($cambios.Count -eq 0 -and $premCambio.Count -eq 0) {
      [void]$lineas.Add("## Resultado: SIN CAMBIOS")
      [void]$lineas.Add("")
      [void]$lineas.Add("Los numeros de " + $mesActual + " son identicos a los que ya estaban publicados. Ya se habian tomado todas las modificaciones del mes: el cierre esta firme.")
    } else {
      [void]$lineas.Add("## Resultado: HUBO CAMBIOS (entraron modificaciones tardias)")
      [void]$lineas.Add("")
      if ($cambios.Count -gt 0) {
        [void]$lineas.Add("### Totales que cambiaron")
        foreach ($c in $cambios) { [void]$lineas.Add($c) }
        [void]$lineas.Add("")
      }
      if ($premCambio.Count -gt 0) {
        [void]$lineas.Add("### Premios de fleteros que cambiaron (OJO, es plata)")
        foreach ($c in $premCambio) { [void]$lineas.Add($c) }
      } else {
        [void]$lineas.Add("Ningun premio de fletero cambio.")
      }
    }
  }
  [void]$lineas.Add("")
  [void]$lineas.Add("Totales definitivos de " + $mesActual + ": efectividad " + $vn["efGeneral"] + "%, carton " + $vn["cartonGeneral"] + "%, " + $vn["repartos"] + " repartos, premios `$" + $vn["premiosTotal"] + ", " + $vn["fleteros"] + " fleteros.")
  [System.IO.File]::WriteAllText((Join-Path $RAIZ ("verificacion-" + $mesActual + ".md")), (($lineas -join "`n")), $utf8)
  Log ("Informe de verificacion escrito: verificacion-" + $mesActual + ".md")
}

# ============================================================================
# 4) Publicacion: la hace el workflow de GitHub Actions (commit de data.js
#    e historial-meses.json). Este script solo deja los archivos listos.
# ============================================================================
Log "================ FIN ================"
