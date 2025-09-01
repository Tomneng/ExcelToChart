package backend.exceltochart.model

import jakarta.persistence.Entity
import jakarta.persistence.Id
import java.time.LocalDateTime
import java.util.UUID

@Entity
class CellStructure(
    @Id val id: UUID = UUID.randomUUID(),

    val cellRole: CellRole,

    val createdAt: LocalDateTime,

) {
}
enum class CellRole{
    HEADER,
    DATA_VALUE,
    CATEGORY,
    SUB_HEADER,
    BLANK,
    ROW_HEADER,
}