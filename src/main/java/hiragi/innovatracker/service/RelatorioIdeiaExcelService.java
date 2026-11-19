package hiragi.innovatracker.service;


import hiragi.innovatracker.dto.IdeiaResponseDTO;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;


import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;


@Slf4j
@Service
public class RelatorioIdeiaExcelService {


    private static final DateTimeFormatter DATETIME_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm");
    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy");


    /**
     * Método principal que coordena a montagem da planilha .xlsx.
     */
    public byte[] gerarRelatorioIdeias(List<IdeiaResponseDTO> ideias, String tituloRelatorio) {
        log.info("Gerando planilha Excel (.xlsx) para {} ideias", ideias.size());


        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {


            Sheet sheet = workbook.createSheet("Relatório de Ideias");


            // 1. Criar repositório de estilos da planilha
            EstilosExcel estilos = criarEstilos(workbook);


            // 2. Executar etapas especialistas de montagem
            int linhaAtual = 0;
            linhaAtual = adicionarBlocoCabecalho(sheet, tituloRelatorio, ideias.size(), estilos, linhaAtual);
            linhaAtual = adicionarTabelaDados(sheet, ideias, estilos, linhaAtual);


            // 3. Ajustar largura automática das colunas (Auto-fit)
            ajustarTamanhoColunas(sheet, 8);


            workbook.write(out);
            return out.toByteArray();


        } catch (IOException e) {
            log.error("Erro ao gerar arquivo Excel", e);
            throw new RuntimeException("Falha na geração do relatório em formato Excel (.xlsx)", e);
        }
    }


    // =========================================================================
    // ESTRUTURA DE SUPORTE E ESTILOS DA PLANILHA
    // =========================================================================


    private record EstilosExcel(
            CellStyle tituloStyle,
            CellStyle subtituloStyle,
            CellStyle headerTableStyle,
            CellStyle dataStyle,
            CellStyle dataZebraStyle,
            CellStyle dataCenterStyle,
            CellStyle dataZebraCenterStyle,
            CellStyle currencyStyle,
            CellStyle currencyZebraStyle,
            CellStyle totalLabelStyle,
            CellStyle totalCurrencyStyle,
            CellStyle statusAprovadoStyle,
            CellStyle statusRejeitadoStyle,
            CellStyle statusProjetoStyle,
            CellStyle statusSubmetidoStyle
    ) {}


    private EstilosExcel criarEstilos(Workbook workbook) {
        DataFormat dataFormat = workbook.createDataFormat();


        // Fontes
        Font fontTitle = workbook.createFont();
        fontTitle.setFontName("Calibri");
        fontTitle.setFontHeightInPoints((short) 16);
        fontTitle.setBold(true);
        fontTitle.setColor(IndexedColors.DARK_BLUE.getIndex());


        Font fontSubtitle = workbook.createFont();
        fontSubtitle.setFontName("Calibri");
        fontSubtitle.setFontHeightInPoints((short) 10);
        fontSubtitle.setItalic(true);
        fontSubtitle.setColor(IndexedColors.GREY_50_PERCENT.getIndex());


        Font fontHeaderTable = workbook.createFont();
        fontHeaderTable.setFontName("Calibri");
        fontHeaderTable.setFontHeightInPoints((short) 11);
        fontHeaderTable.setBold(true);
        fontHeaderTable.setColor(IndexedColors.WHITE.getIndex());


        Font fontBold = workbook.createFont();
        fontBold.setFontName("Calibri");
        fontBold.setBold(true);


        // Estilo Título e Subtítulo
        CellStyle tituloStyle = workbook.createCellStyle();
        tituloStyle.setFont(fontTitle);


        CellStyle subtituloStyle = workbook.createCellStyle();
        subtituloStyle.setFont(fontSubtitle);


        // Estilo Cabeçalho de Tabela
        CellStyle headerTableStyle = workbook.createCellStyle();
        headerTableStyle.setFont(fontHeaderTable);
        headerTableStyle.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
        headerTableStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerTableStyle.setAlignment(HorizontalAlignment.CENTER);
        headerTableStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        configurarBordas(headerTableStyle);


        // Estilos de Dados Básicos (Normal e Zebra)
        CellStyle dataStyle = workbook.createCellStyle();
        configurarBordas(dataStyle);


        CellStyle dataZebraStyle = workbook.createCellStyle();
        dataZebraStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        dataZebraStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        configurarBordas(dataZebraStyle);


        // Alinhamento Centralizado
        CellStyle dataCenterStyle = workbook.createCellStyle();
        dataCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        configurarBordas(dataCenterStyle);


        CellStyle dataZebraCenterStyle = workbook.createCellStyle();
        dataZebraCenterStyle.cloneStyleFrom(dataCenterStyle);
        dataZebraCenterStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        dataZebraCenterStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        // Formatação Monetária (R$ #,##0.00)
        short currencyFormat = dataFormat.getFormat("R$ #,##0.00");


        CellStyle currencyStyle = workbook.createCellStyle();
        currencyStyle.setDataFormat(currencyFormat);
        currencyStyle.setAlignment(HorizontalAlignment.RIGHT);
        configurarBordas(currencyStyle);


        CellStyle currencyZebraStyle = workbook.createCellStyle();
        currencyZebraStyle.cloneStyleFrom(currencyStyle);
        currencyZebraStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        currencyZebraStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        // Estilos da Linha de Totalização
        CellStyle totalLabelStyle = workbook.createCellStyle();
        totalLabelStyle.setFont(fontBold);
        totalLabelStyle.setAlignment(HorizontalAlignment.RIGHT);
        configurarBordas(totalLabelStyle);


        CellStyle totalCurrencyStyle = workbook.createCellStyle();
        totalCurrencyStyle.setFont(fontBold);
        totalCurrencyStyle.setDataFormat(currencyFormat);
        totalCurrencyStyle.setAlignment(HorizontalAlignment.RIGHT);
        totalCurrencyStyle.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        totalCurrencyStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        configurarBordas(totalCurrencyStyle);


        // Estilos com destaque condicional para Status
        CellStyle statusAprovadoStyle = criarEstiloStatus(workbook, IndexedColors.GREEN, IndexedColors.WHITE);
        CellStyle statusRejeitadoStyle = criarEstiloStatus(workbook, IndexedColors.RED, IndexedColors.WHITE);
        CellStyle statusProjetoStyle = criarEstiloStatus(workbook, IndexedColors.ORANGE, IndexedColors.WHITE);
        CellStyle statusSubmetidoStyle = criarEstiloStatus(workbook, IndexedColors.LIGHT_BLUE, IndexedColors.WHITE);


        return new EstilosExcel(
                tituloStyle, subtituloStyle, headerTableStyle, dataStyle, dataZebraStyle,
                dataCenterStyle, dataZebraCenterStyle, currencyStyle, currencyZebraStyle,
                totalLabelStyle, totalCurrencyStyle, statusAprovadoStyle, statusRejeitadoStyle,
                statusProjetoStyle, statusSubmetidoStyle
        );
    }


    private CellStyle criarEstiloStatus(Workbook workbook, IndexedColors bg, IndexedColors fg) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(fg.getIndex());


        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(bg.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        configurarBordas(style);
        return style;
    }


    private void configurarBordas(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }


    // =========================================================================
    // FUNÇÕES DE MONTAGEM DOS BLOCOS
    // =========================================================================


    private int adicionarBlocoCabecalho(Sheet sheet, String titulo, int totalRegistros, EstilosExcel estilos, int startRow) {
        // Linha 0: Título Principal
        Row rowTitle = sheet.createRow(startRow++);
        Cell cellTitle = rowTitle.createCell(0);
        cellTitle.setCellValue(titulo);
        cellTitle.setCellStyle(estilos.tituloStyle());


        // Linha 1: Subtítulo / Metadados de emissão
        Row rowSub = sheet.createRow(startRow++);
        Cell cellSub = rowSub.createCell(0);
        cellSub.setCellValue("InnovaTracker | Emissão: " + LocalDateTime.now().format(DATETIME_FORMATTER) + " | Total de registros: " + totalRegistros);
        cellSub.setCellStyle(estilos.subtituloStyle());


        // Espaçador
        startRow++;
        return startRow;
    }


    private int adicionarTabelaDados(Sheet sheet, List<IdeiaResponseDTO> ideias, EstilosExcel estilos, int startRow) {
        String[] headers = {"ID", "Nome da Ideia", "Autor (Pessoa)", "Categoria", "Orçamento", "Status", "Data Criação", "Resumo"};


        // 1. Linha de Cabeçalho da Tabela
        Row headerRow = sheet.createRow(startRow++);
        headerRow.setHeightInPoints(24);


        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(estilos.headerTableStyle());
        }


        if (ideias.isEmpty()) {
            Row emptyRow = sheet.createRow(startRow++);
            Cell emptyCell = emptyRow.createCell(0);
            emptyCell.setCellValue("Nenhuma ideia encontrada para os filtros aplicados.");
            sheet.addMergedRegion(new CellRangeAddress(startRow - 1, startRow - 1, 0, headers.length - 1));
            return startRow;
        }


        int primeiraLinhaDados = startRow + 1; // 1-based index do Excel para fórmula


        // 2. Preenchimento das Linhas de Dados
        for (int index = 0; index < ideias.size(); index++) {
            IdeiaResponseDTO i = ideias.get(index);
            Row row = sheet.createRow(startRow++);
            boolean zebra = (index % 2 == 1);


            preencherLinhaIdeia(row, i, estilos, zebra);
        }


        int ultimaLinhaDados = startRow; // 1-based index do Excel para fórmula


        // 3. Adicionar Linha de Totalização com Fórmula do Excel (=SUM)
        adicionarLinhaTotalizacao(sheet, startRow++, primeiraLinhaDados, ultimaLinhaDados, estilos);


        return startRow;
    }


    private void preencherLinhaIdeia(Row row, IdeiaResponseDTO i, EstilosExcel estilos, boolean zebra) {
        // Coluna 0: ID
        Cell cellId = row.createCell(0);
        if (i.idtIdeia() != null) cellId.setCellValue(i.idtIdeia());
        cellId.setCellStyle(zebra ? estilos.dataZebraCenterStyle() : estilos.dataCenterStyle());


        // Coluna 1: Nome da Ideia
        Cell cellNome = row.createCell(1);
        cellNome.setCellValue(i.nmeIdeia() != null ? i.nmeIdeia() : "-");
        cellNome.setCellStyle(zebra ? estilos.dataZebraStyle() : estilos.dataStyle());


        // Coluna 2: Autor
        Cell cellPessoa = row.createCell(2);
        cellPessoa.setCellValue(i.nmePessoa() != null ? i.nmePessoa() : "-");
        cellPessoa.setCellStyle(zebra ? estilos.dataZebraStyle() : estilos.dataStyle());


        // Coluna 3: Categoria
        Cell cellCat = row.createCell(3);
        cellCat.setCellValue(i.nmeCategoria() != null ? i.nmeCategoria() : "-");
        cellCat.setCellStyle(zebra ? estilos.dataZebraStyle() : estilos.dataStyle());


        // Coluna 4: Orçamento (Numérico para permitir cálculos no Excel)
        Cell cellOrcamento = row.createCell(4);
        if (i.vlrOrcamentoIdeia() != null) {
            cellOrcamento.setCellValue(i.vlrOrcamentoIdeia().doubleValue());
        } else {
            cellOrcamento.setCellValue(0.0);
        }
        cellOrcamento.setCellStyle(zebra ? estilos.currencyZebraStyle() : estilos.currencyStyle());


        // Coluna 5: Status Formatado com Cor
        Cell cellStatus = row.createCell(5);
        cellStatus.setCellValue(formatarStatusDescricao(i.stsIdeia()));
        cellStatus.setCellStyle(obterEstiloStatus(i.stsIdeia(), estilos));


        // Coluna 6: Data de Criação
        Cell cellData = row.createCell(6);
        cellData.setCellValue(i.dttCriacaoIdeia() != null ? i.dttCriacaoIdeia().format(DATE_FORMATTER) : "-");
        cellData.setCellStyle(zebra ? estilos.dataZebraCenterStyle() : estilos.dataCenterStyle());


        // Coluna 7: Resumo
        Cell cellResumo = row.createCell(7);
        cellResumo.setCellValue(i.txtResumoIdeia() != null ? i.txtResumoIdeia() : "-");
        cellResumo.setCellStyle(zebra ? estilos.dataZebraStyle() : estilos.dataStyle());
    }


    private void adicionarLinhaTotalizacao(Sheet sheet, int rowIndex, int startRowFormula, int endRowFormula, EstilosExcel estilos) {
        Row rowTotal = sheet.createRow(rowIndex);


        // Mesclar células de rótulo (Colunas 0 a 3)
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 3));
        Cell labelCell = rowTotal.createCell(0);
        labelCell.setCellValue("TOTAL DE ORÇAMENTOS:");
        labelCell.setCellStyle(estilos.totalLabelStyle());


        // Aplicar estilo de borda nas células ocultadas da mesclagem
        for (int col = 1; col <= 3; col++) {
            rowTotal.createCell(col).setCellStyle(estilos.totalLabelStyle());
        }


        // Coluna 4: Fórmula =SUM(E{start}:E{end})
        Cell cellSum = rowTotal.createCell(4);
        String formula = String.format("SUM(E%d:E%d)", startRowFormula, endRowFormula);
        cellSum.setCellFormula(formula);
        cellSum.setCellStyle(estilos.totalCurrencyStyle());


        // Preencher bordas restantes
        for (int col = 5; col <= 7; col++) {
            rowTotal.createCell(col).setCellStyle(estilos.totalLabelStyle());
        }
    }


    private String formatarStatusDescricao(String status) {
        if (status == null) return "-";
        return switch (status.toUpperCase()) {
            case "S" -> "SUBMETIDA";
            case "A" -> "APROVADA";
            case "R" -> "REJEITADA";
            case "P" -> "PROJETO";
            default -> status.toUpperCase();
        };
    }


    private CellStyle obterEstiloStatus(String status, EstilosExcel estilos) {
        if (status == null) return estilos.dataCenterStyle();
        return switch (status.toUpperCase()) {
            case "A" -> estilos.statusAprovadoStyle();
            case "R" -> estilos.statusRejeitadoStyle();
            case "P" -> estilos.statusProjetoStyle();
            case "S" -> estilos.statusSubmetidoStyle();
            default -> estilos.dataCenterStyle();
        };
    }


    private void ajustarTamanhoColunas(Sheet sheet, int totalColunas) {
        for (int i = 0; i < totalColunas; i++) {
            sheet.autoSizeColumn(i);
            // Garante uma largura mínima legível em pixels (~12 caracteres)
            int currentWidth = sheet.getColumnWidth(i);
            if (currentWidth < 3600) {
                sheet.setColumnWidth(i, 3600);
            }
        }
    }
}
