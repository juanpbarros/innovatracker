package hiragi.innovatracker.controller;


import hiragi.innovatracker.service.TbIdeiaRelatorioExcelService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.responses.ApiResponses;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;


import java.time.LocalDate;


@RestController
@RequestMapping("/api/ideias/relatorios")
@RequiredArgsConstructor
@Tag(name = "Relatórios de Ideias", description = "Endpoints para geração e download de relatórios em formato Excel (.xlsx)")
public class TbIdeiaRelatorioController {


    private final TbIdeiaRelatorioExcelService ideiaRelatorioExcelService;
    private static final String XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";


    @Operation(summary = "Exportar relatório geral de ideias em Excel", description = "Gera uma planilha .xlsx contendo a listagem completa de ideias e orçamentos.")
    @ApiResponses(value = {
            @ApiResponse(responseCode = "200", description = "Planilha Excel gerada com sucesso"),
            @ApiResponse(responseCode = "500", description = "Erro interno ao gerar a planilha", content = @Content)
    })
    @GetMapping("/excel")
    public ResponseEntity<byte[]> exportarRelatorioGeralExcel() {
        byte[] xlsxBytes = ideiaRelatorioExcelService.gerarRelatorioGeralExcel();
        String filename = "Relatorio_Geral_Ideias_" + LocalDate.now() + ".xlsx";


        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"")
                .contentType(MediaType.parseMediaType(XLSX_MEDIA_TYPE))
                .body(xlsxBytes);
    }


    @Operation(summary = "Exportar relatório de ideias por status em Excel", description = "Gera uma planilha .xlsx filtrada pelo status da ideia ('S', 'A', 'R' ou 'P').")
    @ApiResponses(value = {
            @ApiResponse(responseCode = "200", description = "Planilha Excel gerada com sucesso"),
            @ApiResponse(responseCode = "400", description = "Status informado é inválido", content = @Content)
    })
    @GetMapping("/status/excel")
    public ResponseEntity<byte[]> exportarRelatorioPorStatusExcel(
            @Parameter(description = "Status da ideia ('S', 'A', 'R' ou 'P')", example = "A")
            @RequestParam String status) {


        byte[] xlsxBytes = ideiaRelatorioExcelService.gerarRelatorioPorStatusExcel(status);
        String filename = "Relatorio_Ideias_Status_" + status.toUpperCase() + "_" + LocalDate.now() + ".xlsx";


        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"")
                .contentType(MediaType.parseMediaType(XLSX_MEDIA_TYPE))
                .body(xlsxBytes);
    }


    @Operation(summary = "Exportar relatório de ideias por categoria em Excel", description = "Gera uma planilha .xlsx filtrada por uma categoria específica.")
    @ApiResponses(value = {
            @ApiResponse(responseCode = "200", description = "Planilha Excel gerada com sucesso"),
            @ApiResponse(responseCode = "404", description = "Categoria não encontrada", content = @Content)
    })
    @GetMapping("/categoria/{codCategoria}/excel")
    public ResponseEntity<byte[]> exportarRelatorioPorCategoriaExcel(
            @Parameter(description = "Código da categoria", example = "2")
            @PathVariable Long codCategoria) {


        byte[] xlsxBytes = ideiaRelatorioExcelService.gerarRelatorioPorCategoriaExcel(codCategoria);
        String filename = "Relatorio_Ideias_Categoria_" + codCategoria + "_" + LocalDate.now() + ".xlsx";


        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"")
                .contentType(MediaType.parseMediaType(XLSX_MEDIA_TYPE))
                .body(xlsxBytes);
    }
}
