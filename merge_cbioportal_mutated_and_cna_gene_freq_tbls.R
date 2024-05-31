
library(openxlsx)
library(readxl)
library(dplyr)

xls_file <- "KRAS co-mutated_CNA_Genes.xls"

sheet_names <- excel_sheets(xls_file)

df1 <- read_excel(xls_file, sheet = sheet_names[1]) %>%
  select(Gene, Freq) %>%
  mutate(Freq = as.character(Freq)) %>%
  mutate(Freq = case_when(
    grepl("<", Freq) ~ Freq,
    TRUE ~ as.character(round(as.numeric(Freq), 3))
  ))

df2 <- read_excel(xls_file, sheet = sheet_names[2]) %>%
  select(Gene, Freq)

# harmonize dups in df2
genes_df2 <- unique(df2$Gene)
df2_h <- do.call(rbind, lapply(genes_df2, function(g) {
  
  tmp <- df2 %>% 
    filter(Gene == g)
  
  if (nrow(tmp) == 1) {
    
    if (!grepl("<", tmp$Freq)) {
      tmp <- tmp %>%
        mutate(Freq = round(as.numeric(Freq), 3))
    }
    return (tmp)
    
  } else {
    
    if (all(grepl("<", tmp$Freq))) {
      return (tmp[1, ])
    }
    
    tmp <- tmp %>%
      filter(!grepl("<", Freq)) %>%
      mutate(Freq = as.numeric(Freq))
    
    
    return (data.frame(
      Gene = tmp$Gene[1], 
      Freq = round(sum(tmp$Freq), 3)
    ))
  }
}))


merged_df <- df1 %>%
  left_join(df2_h, by = "Gene", suffix = c(".Mut", ".CNA")) %>%
  mutate(
    Freq.Mut = as.character(Freq.Mut), 
    Freq.CNA = as.character(Freq.CNA)
  ) %>%
  mutate(
    Freq.Mut = 
      case_when(
        grepl("<", Freq.Mut) | is.na(Freq.Mut) ~ Freq.Mut,
        TRUE ~ paste0(round(as.numeric(Freq.Mut) * 100, 1), "%")
      ), 
    Freq.CNA = 
      case_when(
        grepl("<", Freq.CNA) | is.na(Freq.CNA) ~ Freq.CNA,
        TRUE ~ paste0(round(as.numeric(Freq.CNA) * 100, 1), "%")
      )
  )

write.xlsx(merged_df, "KRAS_merged.xlsx", rowNames = FALSE)

