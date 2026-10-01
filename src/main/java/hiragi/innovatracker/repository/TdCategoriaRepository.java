package hiragi.innovatracker.repository;

import hiragi.innovatracker.model.TdCategoria;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface TdCategoriaRepository extends JpaRepository<TdCategoria, Long> {

    Optional<TdCategoria> findBySglCategoriaIgnoreCase(String sglCategoria);

    boolean existsBySglCategoriaIgnoreCase(String sglCategoria);

    Optional<TdCategoria> findByNmeCategoriaIgnoreCase(String nmeCategoria);

    boolean existsByNmeCategoriaIgnoreCase(String nmeCategoria);

    List<TdCategoria> findByNmeCategoriaContainingIgnoreCaseOrderByNmeCategoriaAsc(
            String nmeCategoria
    );
}