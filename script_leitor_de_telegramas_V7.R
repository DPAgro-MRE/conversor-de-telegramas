# inspirado em https://learningactors.com/how-to-extract-and-clean-data-from-pdf-files-in-r/

# OBS: Por enquanto funciona apenas com TELEGRAMAS
# Requisitos:
# 1) Gravar arquivo pdftotext.exe no working directory
# 2) Gravar CSV com lista de países no working directory (exportar de https://docs.google.com/spreadsheets/d/1jwJT_czC6dQBvUIFEj9f8X7rHsUOIfWgarlguAhodq0/edit?usp=sharing)

library(tm)
library(tidyverse)
library(lubridate)
library(here)
library(writexl)


options(tibble.print_max = 20) 
options(tibble.width = Inf) 

# carrega csv de países e exclui "Brasil"
tabela_paises <- read_csv("Knack - planilhas de apoio - lista_paises.csv") %>% 
  filter(.$pais_detecta != "Brasil") %>% 
  mutate(pais_detecta = tolower(pais_detecta)) %>% 
  mutate(pais_detecta = str_c(pais_detecta, "."))

paises_detectar <- paste(tabela_paises$pais_detecta, collapse = "|")

arquivo_entrada <- file.choose()

read <- readPDF(engine = "xpdf", control = list(text = "-simple"))
document <- Corpus(URISource(arquivo_entrada), readerControl = list(reader = read))

doc <- content(document[[1]]) %>%
  str_c(collapse = "\r\n") %>% # junta todos os registros (linhas)
  strsplit("\f", Inf) %>% # separa novamente registros, pelo caractere de quebra de página 
  flatten_chr() %>%
  enframe %>%

  # captura valores cabeçalho
    mutate(posto_cabecalho = str_extract(value, regex("De: .*"))) %>%
    mutate(posto_cabecalho = str_remove(posto_cabecalho, "De: ")) %>%
    mutate(posto_cabecalho = str_sub(posto_cabecalho, 1, 25)) %>%
    mutate(posto_cabecalho = str_trim(posto_cabecalho)) %>%
  
    mutate(numero = str_extract(value, regex("N.°: \\d*"))) %>%
    mutate(numero = str_remove(numero, "N.°: ")) %>% 
    mutate(numero = as.numeric(numero)) %>%
  
    mutate(data_hora_entrada = str_extract(value, regex("Recebido em: \\d{2}/\\d{2}/\\d{4} \\d{2}:\\d{2}:\\d{2}|Expedido em: \\d{2}/\\d{2}/\\d{4} \\d{2}:\\d{2}:\\d{2}"))) %>%
    mutate(data_hora_entrada = str_remove(data_hora_entrada, "Recebido em: |Expedido em: ")) %>%
    mutate(data_documento = dmy(str_sub(data_hora_entrada, end = 10))) %>%
    mutate(ano = year(data_documento)) %>% 
         #checar acima, está resultando so dois digitos
    mutate(data_hora_entrada = dmy_hms(data_hora_entrada)) %>%

  
  # retira rodapé
    mutate(value = str_replace_all(value, regex("\\s*Distribuído em: \\d{2}/\\d{2}/\\d{4} \\d{2}:\\d{2}:\\d{2}\\s*Impresso em: \\d{2}/\\d{2}/\\d{4} - \\d{2}:\\d{2}\\r?\\n?"), "\r\n")) %>%
  
  # retira cabeçalho
    mutate(value = str_remove_all(value, regex("De: .*Página \\d*/\\d*\\r\\nCARAT=\\w*\\r\\n\\s*Recebido em: \\d{2}/\\d{2}/\\d{4} \\d{2}:\\d{2}:\\d{2} N.°: \\d*\\r\\n\\s*"))) %>%
    mutate(value = str_remove_all(value, regex("\\s*Página \\d*/\\d*\\r\\n\\r\\nDe: .*Recebido em: \\d{2}/\\d{2}/\\d{4} \\d{2}:\\d{2}:\\d{2} N.°: \\d*\\r\\nCARAT=\\w*"))) %>%
    mutate(value = str_remove_all(value, regex("\\s*Código de autenticação: .*(\\r\\n){0,2}"))) %>%
  
  # funde páginas
    group_by(posto_cabecalho, numero) %>%
    mutate(value = str_c(value, collapse = "\r\n")) %>%
    distinct(posto_cabecalho, numero, .keep_all = TRUE) %>% 
    ungroup() %>%  
  
  # retira espaços à esquerda
    mutate(value = str_replace_all(value, "\\r\\n   ", "\r\n")) %>%
    mutate(value = str_remove_all(value, "^   ")) %>%

  #captura valores início telegrama
    mutate(tipo = "TEL") %>%
  
    mutate(posto = str_extract(value, regex("De .* para Exteriores"))) %>%
    mutate(posto = str_remove(posto, regex("De "))) %>%
    mutate(posto = str_remove(posto, regex(" para Exteriores"))) %>%

    mutate(carater = str_extract(value, regex("CARAT=\\w*"))) %>%
    mutate(carater = str_remove_all(carater, "CARAT=")) %>%
    filter(carater == "Ostensivo") %>% 
 
    mutate(prioridade = str_extract(value, regex("PRIOR=\\w*", ignore_case = TRUE))) %>%
    mutate(prioridade = str_remove_all(prioridade, "PRIOR=")) %>%

    mutate(distribuicao = str_extract(value, regex("DISTR=[A-Z a-z /]+"))) %>%
    mutate(distribuicao = str_remove(distribuicao, "DISTR=")) %>% 

    mutate(primdistribuicao = str_extract(distribuicao, regex("[A-Z a-z]+"))) %>%
  
  
    mutate(redistribuicao = str_extract(value, regex("Nota da DCA: Redistribuído para .* em *"))) %>%
    mutate(redistribuicao = str_remove(redistribuicao, regex("Nota da DCA: Redistribuído para "))) %>%
    mutate(redistribuicao = str_remove(redistribuicao, regex(" em *"))) %>% 
  
    mutate(primredistribuicao = str_extract(redistribuicao, regex("[A-Z a-z]+"))) %>%
    

    mutate(indice = str_extract(value, regex("//\\r\\n.*\\r\\n//", dotall = TRUE)))  %>%
    mutate(indice = str_remove(indice, regex("^//(\\r\\n)*"))) %>%
    mutate(indice = str_remove(indice, regex("(\\r\\n)*//$"))) %>%
    mutate(indice = str_replace_all(indice, "\\r\\n", " ")) %>%
    mutate(indice = str_remove(indice, regex("^\\s*"))) %>% # rever depois pq precisa
  
     # detecta primeiro país/OI dentro do índice
      mutate(pais_detecta = str_extract(indice, regex(paises_detectar, ignore_case = TRUE))) %>%
  
    mutate(refdoc = str_extract(value, regex("REF/ADIT=.*")))  %>%
    mutate(refdoc = str_remove(refdoc, "REF/ADIT=")) %>% 
    mutate(refdoc = str_replace_all(refdoc, "(TEL [0-9]+|DET [0-9]+) ([0-9]{4})", "\\1/\\2/<posto>")) %>%
    mutate(refdoc = str_replace_all(refdoc, "(TEL [0-9]+|DET [0-9]+),", "\\1/<ano>/<posto>,")) %>%
    mutate(refdoc = str_replace_all(refdoc, "(TEL [0-9]+|DET [0-9]+)$", "\\1/<ano>/<posto>")) %>%
    mutate(refdoc = str_replace_all(refdoc, "<ano>", as.character(.$ano))) %>%
    mutate(refdoc = str_replace_all(refdoc, "<posto>", .$posto)) %>%
 
   
  # retira quebras de linha do corpo do telegrama
    mutate(corpo = str_extract(value, regex("Nr. \\d+.*", dotall = TRUE)))  %>%
    mutate(corpo = str_remove(corpo, regex("Nr. \\d+\\r?\\n?"))) %>%
    mutate(corpo = str_replace_all(corpo, "\\r\\n\\r\\n", "<pArAg>")) %>%
    mutate(corpo = str_replace_all(corpo, "\\r\\n", " ")) %>%
    mutate(corpo = str_replace_all(corpo, "<pArAg>", "\r\n\r\n")) %>%
  
    mutate(corpo = str_remove(corpo, regex("Retransmissão automática.*\\r?\\n?"))) %>%
    mutate(corpo = str_remove(corpo, regex("Retransmitido via clic.*\\r?\\n?"))) %>%
  
  # extrai Cumpre Instruções
    mutate(cumpre_instruções = if_else(str_detect(corpo, regex("(cumpre |cumpri )(instrução|instruções)",
                                                               ignore_case = TRUE)), "Sim", "Não")) %>%
  
  # extrai resumo do corpo
    mutate(corpo = str_replace_all(corpo, "RESUMO=(\\r\\n)*", "RESUMO=")) %>%
    mutate(resumo = str_extract(corpo, regex("RESUMO=\\s*.*"))) %>%
    mutate(resumo = str_remove(resumo, "RESUMO=")) %>%
    mutate(corpo = str_remove(corpo, regex("RESUMO=.*"))) %>%
    mutate(corpo = str_remove(corpo, regex("(\\r\\n)*"))) %>%
  
  # extrai Processos SEI
    mutate(processos_sei = str_extract(corpo, regex("\\d{5}\\.\\d{6}/\\d{4}-\\d{2}"))) %>% 

  # altera nome de país/OI detectado para nome no sistema
    mutate(pais_detecta = tolower(pais_detecta)) %>% 
    left_join(tabela_paises) %>% 
    
  # inclui chave
    mutate(documento = paste0("TEL ", .$numero, "/", .$ano, "/", .$posto)) %>%

  # reorganiza tabela
    select("data_hora_entrada" = data_hora_entrada,
           "data_documento" = data_documento,
           "tipo_documento" = tipo,
           "numero" = numero,
           "ano" = ano,
           "remetente" = posto,
           "documento" = documento,
           "indice" = indice,
           "prioridade" = prioridade,
           "carater" = carater,
           "distribuicao" = distribuicao,
           primdistribuicao,
           "redistribuicao" = redistribuicao, primredistribuicao,
           "refdoc" = refdoc,
           "processos_sei" = processos_sei,
           "teor" = value,
           "corpo" = corpo,
           "resumo" = resumo,
           "paises_ois" = pais_final,
           "pasta" = nome_pasta,
           "apenas_cumpre_instrucoes" = cumpre_instruções) %>%

  # filtra tels de interesse
    filter(primdistribuicao == "DPAGRO" | primdistribuicao == "DPA I" | primdistribuicao == "DPagro" | primdistribuicao == "DPAgro"
           | primredistribuicao == "DPAGRO")


# grava o arquivo de saída na pasta ./saida
write_csv(doc, str_c("./saida/", str_replace(basename(arquivo_entrada), "pdf", "csv")))

caminho_excel <- here("saida", str_replace(basename(arquivo_entrada), "pdf", "xlsx"))
write_xlsx(doc, caminho_excel)




# Falta: redistribuição (Nota da DCA: Redistribuído para DIOEC/DACESS/DPA I/DDF/DNS/DAP em 18/03/2021)
