package amz.billing.Trips.entities;

import lombok.*;
import org.springframework.data.mongodb.core.mapping.Document;

import java.time.LocalDateTime;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@EqualsAndHashCode
@Builder
public class Stop {

    private String stopName;
    private LocalDateTime stopYardArrival;
    private LocalDateTime stopYardDeparture;
}
