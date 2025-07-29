package amz.billing.Trips.entities;

import lombok.*;
import org.springframework.data.annotation.Id;
import org.springframework.data.mongodb.core.index.Indexed;
import org.springframework.data.mongodb.core.mapping.Document;
import org.springframework.data.mongodb.core.mapping.Field;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@EqualsAndHashCode
@Builder
@Document(collection = "company")
public class Company {

    @Id
    private String id;

    @Field(name = "key")
    @Indexed(unique = true)
    private String key;

    @Field(name = "truck_id_list")
    private List<String> truckIdList;

    @Field(name = "company_name")
    private String fullNameCompany;

    @Field(name = "discount")
    private double discount;

    @Field(name = "discount_aveka")
    private double discountAveka;
}
