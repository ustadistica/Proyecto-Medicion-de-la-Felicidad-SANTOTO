# ============================================================
#        ANÁLISIS DE ESCALA DE BIENESTAR SUBJETIVO
# ============================================================

# --- Paquetes ---
library(readxl)
library(dplyr)
library(tidyr)
library(psych)
library(polycor)
library(openxlsx)
library(janitor)
library(ggplot2)

# ============================================================
# 1. LECTURA DE ARCHIVOS
# ============================================================

df <- read_excel("Insumos/Formulario de la escala de bienestar subjetivo(1-280).xlsx")

tab <- read_excel(
  "Insumos/TABULACIÓN DE RESULTADOS PRUEBA PILOTO.xlsx",
  range = "A1:K19"
) %>%
  mutate(
    categoria = case_when(
      row_number() %in% 1:2   ~ "AUTOACEPTACIÓN",
      row_number() %in% 3:5   ~ "CRECIMIENTO PERSONAL",
      row_number() %in% 6:10  ~ "PROPÓSITO DE VIDA",
      row_number() %in% 11:13 ~ "DOMINIO DEL ENTORNO",
      row_number() %in% 14:16 ~ "RELACIONES POSITIVAS",
      row_number() %in% 17:18 ~ "AUTONOMÍA"
    )
  ) %>%
  select(categoria, everything(), -2)

# ============================================================
# 2. LIMPIEZA Y DEPURACIÓN DE DATOS
# ============================================================

items <- df[, 28:ncol(df)]
it_cols <- paste0("it", 1:length(items))
colnames(items) <- it_cols
str(items)

faltantes_items <- items %>%
  summarise(across(everything(), ~ sum(is.na(.)))) %>%
  pivot_longer(everything(), names_to = "item", values_to = "n_faltantes") %>%
  arrange(desc(n_faltantes))

print(faltantes_items)

df_limpio <- df[complete.cases(items), ]

# write.xlsx(df_limpio, "base_limpia.xlsx", overwrite = TRUE)

# ============================================================
# 3. DETECCIÓN DE ÍTEMS INVERTIDOS
# ============================================================

analizar_grupo <- function(df, items, k = 6) {
  sub <- df[, 27 + items, drop = FALSE]
  colnames(sub) <- paste0("item_", items)
  
  cor_original <- cor(sub, use = "pairwise.complete.obs")
  cor_prom_original <- rowMeans(cor_original, na.rm = TRUE)
  
  cor_prom_invertido <- sapply(seq_along(sub), function(i) {
    sub_i <- sub
    sub_i[[i]] <- (k + 1) - sub_i[[i]]
    cor_mat_i <- cor(sub_i, use = "pairwise.complete.obs")
    rowMeans(cor_mat_i, na.rm = TRUE)[i]
  })
  
  data.frame(
    item = paste0("item_", items),
    cor_prom_original = round(cor_prom_original, 3),
    cor_prom_invertido = round(cor_prom_invertido, 3)
  )
}

analizar_grupo(df_limpio, c(7, 17))      # Autoaceptación
analizar_grupo(df_limpio, c(21, 27, 28)) # Crecimiento personal
analizar_grupo(df_limpio, c(6, 11, 15, 16, 20)) # Propósito de vida
analizar_grupo(df_limpio, c(10, 14, 29)) # Dominio del entorno
analizar_grupo(df_limpio, c(2, 12, 25))  # Relaciones positivas
analizar_grupo(df_limpio, c(3, 18))      # Autonomía

# Invertir ítem identificado
df_limpio$`A menudo me siento solo porque tengo pocos amigos íntimos con quienes compartir mis preocupaciones.` <-
  7 - df_limpio$`A menudo me siento solo porque tengo pocos amigos íntimos con quienes compartir mis preocupaciones.`

# ============================================================
# 4. MEDIDAS DE TENDENCIA CENTRAL Y DISPERSIÓN
# ============================================================

items <- df_limpio[, 28:ncol(df_limpio)]

resumen_items <- items %>%
  summarise(across(everything(),
                   list(
                     media = ~ mean(.x, na.rm = TRUE),
                     mediana = ~ median(.x, na.rm = TRUE),
                     sd = ~ sd(.x, na.rm = TRUE),
                     min = ~ min(.x, na.rm = TRUE),
                     max = ~ max(.x, na.rm = TRUE)
                   ),
                   .names = "{.col}_{.fn}"))

resumen_largo <- resumen_items %>%
  pivot_longer(everything(), names_to = c("item", ".value"), names_sep = "_")

head(resumen_largo)

# ============================================================
# 5. DISTRIBUCIÓN DE FRECUENCIAS POR ÍTEM
# ============================================================

frecuencias_items <- items %>%
  pivot_longer(everything(), names_to = "item", values_to = "valor") %>%
  group_by(item, valor) %>%
  summarise(frecuencia = n(), .groups = "drop_last") %>%
  mutate(porcentaje = round(100 * frecuencia / sum(frecuencia), 2)) %>%
  arrange(item, valor)

print(head(frecuencias_items, 12))

for (i in unique(frecuencias_items$item)) {
  p <- frecuencias_items %>%
    filter(item == i) %>%
    ggplot(aes(x = factor(valor), y = porcentaje, fill = factor(valor))) +
    geom_bar(stat = "identity", color = "black") +
    geom_text(aes(label = paste0(porcentaje, "%")), vjust = -0.3, size = 3) +
    scale_fill_brewer(palette = "Blues") +
    labs(title = paste("Distribución de frecuencias -", i),
         x = "Valor de respuesta (1 a 6)", y = "Porcentaje (%)") +
    theme_minimal()
  print(p)
}

# ============================================================
# 6. MÉTRICA GLOBAL Y RESÚMENES AGRUPADOS
# ============================================================

items_num <- items %>% mutate(across(everything(), as.numeric))

df_limpio <- df_limpio %>%
  mutate(metrica_global = rowSums(items_num, na.rm = TRUE))

# Resumen General
resumen_general <- df_limpio %>%
  summarise(
    total_encuestados = n(),
    suma_total = sum(metrica_global, na.rm = TRUE),
    promedio_general = mean(metrica_global, na.rm = TRUE),
    desviacion_estandar = sd(metrica_global, na.rm = TRUE)
  )

# Resumen Modalidad  
resumen_modalidad <- df_limpio %>%
  group_by(Modalidad) %>%
  summarise(
    promedio_modalidad = mean(metrica_global, na.rm = TRUE),
    n = n(),
    suma_total = sum(metrica_global, na.rm = TRUE)
  )

# Resumen Nivel de formación
resumen_nivel <- df_limpio %>%
  group_by(`Nivel de formación`) %>%
  summarise(
    promedio_nivel = mean(metrica_global, na.rm = TRUE),
    n = n(),
    suma_total = sum(metrica_global, na.rm = TRUE)
  )

# Resumen Género 
resumen_genero <- df_limpio %>%
  group_by(Género) %>%
  summarise(
    promedio_genero = mean(metrica_global, na.rm = TRUE),
    n = n(),
    suma_total = sum(metrica_global, na.rm = TRUE)
  )

# Resumen Estrato Socioeconomico
resumen_estrato <- df_limpio %>%
  group_by(`Estrato socioeconómico`) %>%
  summarise(
    promedio_estrato = mean(metrica_global, na.rm = TRUE),
    n = n(),
    suma_total = sum(metrica_global, na.rm = TRUE)
  )

cat("\nResumen general:\n"); print(resumen_general)
cat("\nResumen por modalidad:\n"); print(resumen_modalidad)
cat("\nResumen por nivel de formación:\n"); print(resumen_nivel)
cat("\nResumen por género:\n"); print(resumen_genero)
cat("\nResumen por estrato socioeconómico:\n"); print(resumen_estrato)

# ============================================================
# 7. DESCRIPTIVOS DE LAS VARIABLES
# ============================================================

describe(items)

# ============================================================
# 8. MATRIZ POLICORICA Y ANALISIS PARALELO
# ============================================================

pc <- hetcor(items)$correlations ; pc
fa.parallel(pc, n.obs = nrow(items), fm = "minres")
efa1 <- fa(pc, nfactors = 1, fm = "minres")

# ============================================================
# 6. FIABILIDAD
# ============================================================

alpha_ord <- alpha(pc, n.obs = nrow(df), check.keys = TRUE)
omega_ord <- omega(pc, n.obs = nrow(df))

# ============================================================
# 10. DISCRIMINACIÓN
# ============================================================

alpha_raw <- alpha(items, check.keys = TRUE)
rit <- alpha_raw$item.stats$r.drop
alpha_if_del <- alpha_raw$alpha.drop$raw_alpha

# ============================================================
# 11. DIFICULTAD DE LAS PREGUNTAS
# ============================================================

item_means <- sapply(items, mean)
p_star <- item_means / k
difficulty <- 1 - p_star

# ============================================================
# SCORE TOTAL Y SEM
# ============================================================

total <- rowMeans(items, na.rm = TRUE) * length(it_cols)
reliab <- alpha_ord$total$raw_alpha # o omega_ord$omega.tot

SEM <- sd(total, na.rm = TRUE) * sqrt(1 - reliab) 
