# Paquetes necesarios
#install.packages(c("readxl","dplyr", "sampling", "writexl"))  # <- descomenta si no los tienes
library(readxl)
library(dplyr)
library(sampling)
library(writexl)

# =========================================================
# 1) Selecciona los archivos (saldrá un cuadro de diálogo)
# =========================================================
cat("Selecciona el archivo de Bucaramanga...\n")
path_bga <- file.choose()

cat("Selecciona el archivo de Tunja...\n")
path_tun <- file.choose()

cat("Selecciona el archivo de Villavicencio...\n")
path_vil <- file.choose()

# =========================================================
# 2) Bucaramanga: tiene una sola hoja -> tomamos esa
# =========================================================
sheets_bga <- excel_sheets(path_bga)
sheet_bga  <- sheets_bga[1]

bga <- read_excel(path_bga, sheet = sheet_bga) 

# =========================================================
# 3) Tunja: usar SOLO la hoja "exportar hoja de trabajo"
#    (ignorando "Secuelas")
# =========================================================
sheets_tun <- excel_sheets(path_tun)
# Buscar por nombre (insensible a mayúsculas/espacios)
target_tun <- grepl("exportar\\s*hoja\\s*de\\s*trabajo",
                    tolower(sheets_tun))
if (!any(target_tun)) {
  message("⚠️ No se encontró 'exportar hoja de trabajo' en Tunja; se usa la primera hoja.")
  sheet_tun <- sheets_tun[1]
} else {
  sheet_tun <- sheets_tun[which(target_tun)[1]]
}

tun <- read_excel(path_tun, sheet = sheet_tun)
tun["SEDE"] <- "Tunja"

# =========================================================
# 4) Villavicencio: leer 2 hojas -> pregrado y pos/postgrado
# =========================================================
sheets_vil <- excel_sheets(path_vil)

# Buscar nombres típicos (pre / pos|post), insensible a mayúsculas
pre_idx <- which(grepl("^\\s*pre", tolower(sheets_vil)))
pos_idx <- which(grepl("pos|post", tolower(sheets_vil)))

# Pequeños respaldos por si los nombres no coinciden exactamente
if (length(pre_idx) == 0)  pre_idx <- 1
if (length(pos_idx) == 0)  pos_idx <- setdiff(1:min(2, length(sheets_vil)), pre_idx)[1]

sheet_pre <- sheets_vil[pre_idx[1]]
sheet_pos <- sheets_vil[pos_idx[1]]

vil_pre <- read_excel(path_vil, sheet = sheet_pre) 
vil_pre["SEDE"] <- "Villavicencio"
vil_pos <- read_excel(path_vil, sheet = sheet_pos) 
vil_pos["SEDE"] <- "Villavicencio"
vil_pos["MODALIDAD"] <- "Posgrado"

# ===========================================================
#
#============================================================

bga= rename(bga,
              SEDE = `Seccional`,
              PROGRAMA  = `Programa`,
              PERÍODO =`Período`,
              NOMBRE = `Nombre`,
              MODALIDAD =`Modalidad`,
              NUM_IDENTIFICACION = `Identificación`,
              ID_TIPO_DOCUMENTO=`Tipo identificación`,
              GÉNERO=`Género`,
              COD_UNIDAD=`Cód. Programa`,
              DIVISION=`División`,
              SEMESTRE=`Semestre`,
              DIR_EMAIL =`Correo electrónico institucional`,
              DIR_EMAIL_PER=`Correo electrónico personal`)

tun= rename(tun,
            PERÍODO =`PRDO_MATRICULA`,
            NOMBRE = `NOM_LARGO`,
            GÉNERO=`GEN_TERCERO`,
            SEMESTRE=`NUM_NIV_CURSA`)


vil_pre= rename(vil_pre,
            PROGRAMA  = `Programa`,
            NOMBRE = `Nombre`,
            PERÍODO =`Período`,
            COD_PENSUM =`Cód. Pensum`,
            NUM_IDENTIFICACION = `Identificación`,
            ID_TIPO_DOCUMENTO=`Tipo identificación`,
            GÉNERO=`Género`,
            COD_UNIDAD=`Cód. Programa`,
            COD_ALUMNO=`Documento`,
            MODALIDAD =`Modalidad`,
            SEMESTRE=`Nivel cursa`,
            DIR_EMAIL =`Correo electrónico institucional`)

vil_pos = rename(vil_pos,
                PROGRAMA  = `Programa`,
                NOMBRE = `Nombre`,
                PERÍODO =`Período`,
                NUM_IDENTIFICACION = `Identificación`,
                ID_TIPO_DOCUMENTO=`Tipo identificación`,
                GÉNERO=`Género`,
                COD_UNIDAD=`Cód. Programa`,
                COD_ALUMNO=`Código del alumno`,
                SEMESTRE=`Nivel cursa`,
                DIR_EMAIL =`Correo electrónico institucional`)


vil <- rbind.data.frame(vil_pos,vil_pre[,-c(7)])

# ===========================================================
#
#============================================================
BGA = dplyr::select(bga, "SEDE", "PERÍODO", "ID_TIPO_DOCUMENTO", "NUM_IDENTIFICACION","NOMBRE","GÉNERO","COD_UNIDAD","PROGRAMA","SEMESTRE","DIR_EMAIL")

TUN = dplyr::select(tun, "SEDE", "PERÍODO", "ID_TIPO_DOCUMENTO", "NUM_IDENTIFICACION","NOMBRE","GÉNERO","COD_UNIDAD","PROGRAMA","SEMESTRE","DIR_EMAIL")

VIL_PRE = dplyr::select(vil_pre, "SEDE","PERÍODO", "ID_TIPO_DOCUMENTO", "NUM_IDENTIFICACION","NOMBRE","GÉNERO","COD_UNIDAD","PROGRAMA","SEMESTRE","DIR_EMAIL")

VIL_POS =dplyr::select(vil_pos, "SEDE", "PERÍODO", "ID_TIPO_DOCUMENTO", "NUM_IDENTIFICACION","NOMBRE","GÉNERO","COD_UNIDAD","PROGRAMA","SEMESTRE","DIR_EMAIL")

# =========================================================
# 5) Unir las CUATRO bases
#    (Bucaramanga 1 + Tunja 1 + Villavicencio 2)
# =========================================================

datos_union <- bind_rows(BGA, TUN, VIL_PRE, VIL_POS)
dim(datos_union)

# Vista rápida
cat("\nResumen por sede y nivel:\n")
print(datos_union %>% count(PROGRAMA,SEMESTRE, name = "N"))

# =================== Muestreo ===================

df <- datos_union               #  Base de datos unida
N <- nrow(df)
n<- ss4p(N, P= 0.5, error = "me", delta = 0.04)
st <- c("SEDE")        # variables de estratificación
nh <- round(prop.table(table(df$SEDE))*n,0)
df <- df %>%
  group_by(SEDE) %>%
  mutate(
    Nh = n(),
    nh = nh[match(unique(!!sym(st)), names(nh))],
    pik = nh / Nh
  )

set.seed(20251023)
muestra_ids <- strata(
  data = df,
  stratanames = st,
  size = nh,
  method = "systematic",
  pik = df$pik
)

# Obtener datos finales
muestra_final <- getdata(df, muestra_ids)
muestra <- muestra_final %>% select(ID_TIPO_DOCUMENTO, NUM_IDENTIFICACION, SEDE, 
                                          NOMBRE, COD_UNIDAD, PROGRAMA, SEMESTRE, DIR_EMAIL)
# Verificar muestreo
table(muestra$SEDE)

# Muestra
salida <- file.choose()
write_xlsx(muestra, salida)







