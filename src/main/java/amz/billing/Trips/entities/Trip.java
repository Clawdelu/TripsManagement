package amz.billing.Trips.entities;

import amz.billing.Trips.enums.Status;
import lombok.*;
import org.springframework.data.annotation.Id;
import org.springframework.data.mongodb.core.index.Indexed;
import org.springframework.data.mongodb.core.mapping.Document;
import org.springframework.data.mongodb.core.mapping.Field;

import java.util.List;

@Getter
@Setter
@EqualsAndHashCode
@NoArgsConstructor
@AllArgsConstructor
@Builder
@Document(collection = "trips")
public class Trip {
    @Id
    private String id;

    @Field(name = "vrid")
    @Indexed(unique = true)
    private String vrid;

    @Field(name = "price")
    private Double price;

    @Field(name = "status")
    private Status status;

    @Field(name = "drivers")
    private List<String> driverList;

    @Field(name = "trailer_ID")
    private String trailerID;

    @Field(name = "vehicle_ID")
    private String vehicleID;

    @Field(name = "stops")
    private List<Stop> stopList;

    @Field(name = "total_distance")
    private Double totalDistance;

    @Field(name = "transit_operator_type")
    private String transitOperatorType;

}
