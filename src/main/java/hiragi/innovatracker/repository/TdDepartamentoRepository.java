package hiragi.innovatracker.repository;

import hiragi.innovatracker.model.TdDepartamento;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface TdDepartamentoRepository extends JpaRepository<TdDepartamento, Long> {

    Optional<TdDepartamento> findBySglDepartamentoIgnoreCase(String sglDepartamento);

    boolean existsBySglDepartamentoIgnoreCase(String sglDepartamento);

    Optional<TdDepartamento> findByNmeDepartamentoIgnoreCase(String nmeDepartamento);

    boolean existsByNmeDepartamentoIgnoreCase(String nmeDepartamento);

    List<TdDepartamento> findByNmeDepartamentoContainingIgnoreCaseOrderByNmeDepartamentoAsc(
            String nmeDepartamento
    );
}