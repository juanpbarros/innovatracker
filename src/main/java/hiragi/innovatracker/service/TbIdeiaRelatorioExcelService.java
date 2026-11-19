package hiragi.innovatracker.service;


import hiragi.innovatracker.dto.IdeiaResponseDTO;
import hiragi.innovatracker.repository.TbIdeiaRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;


import java.util.List;


@Slf4j
@Service
@RequiredArgsConstructor
public class TbIdeiaRelatorioExcelService {


    private final TbIdeiaRepository ideiaRepository;
    private final RelatorioIdeiaExcelService relatorioIdeiaExcelService;


    @Transactional(readOnly = true)
    public byte[] gerarRelatorioGeralExcel() {
        log.info("Buscando todos os registros de ideias para o relatório Excel...");


        List<IdeiaResponseDTO> ideias = ideiaRepository.findAll(Sort.by(Sort.Direction.ASC, "nmeIdeia"))
                .stream()
                .map(IdeiaResponseDTO::from)
                .toList();


        return relatorioIdeiaExcelService.gerarRelatorioIdeias(ideias, "Relatório Geral de Ideias e Orçamentos");
    }


    @Transactional(readOnly = true)
    public byte[] gerarRelatorioPorStatusExcel(String status) {
        log.info("Buscando ideias com status='{}' para relatório Excel...", status);


        List<IdeiaResponseDTO> ideias = ideiaRepository
                .findByStsIdeiaOrderByNmeIdeiaAsc(status)
                .stream()
                .map(IdeiaResponseDTO::from)
                .toList();


        String titulo = "Relatório de Ideias - Status: " + status.toUpperCase();
        return relatorioIdeiaExcelService.gerarRelatorioIdeias(ideias, titulo);
    }


    @Transactional(readOnly = true)
    public byte[] gerarRelatorioPorCategoriaExcel(Long codCategoria) {
        log.info("Buscando ideias da categoria ID={} para relatório Excel...", codCategoria);


        List<IdeiaResponseDTO> ideias = ideiaRepository
                .findByTdCategoria_IdtCategoriaOrderByNmeIdeiaAsc(codCategoria)
                .stream()
                .map(IdeiaResponseDTO::from)
                .toList();


        String titulo = "Relatório de Ideias - Categoria " + codCategoria;
        return relatorioIdeiaExcelService.gerarRelatorioIdeias(ideias, titulo);
    }
}
