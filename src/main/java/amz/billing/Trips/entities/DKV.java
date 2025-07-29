package amz.billing.Trips.entities;

import lombok.*;
import org.springframework.data.annotation.Id;
import org.springframework.data.mongodb.core.mapping.Document;
import org.springframework.data.mongodb.core.mapping.Field;

import java.time.LocalDate;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@EqualsAndHashCode
@Builder
@Document(collection = "DKV")
public class DKV {

    @Id
    private String id;

    @Field(name = "customer_id")
    private String customerId;

    @Field(name = "customer_name")
    private String customerName;

    @Field(name = "invoice_date")
    private LocalDate invoiceDate;

    @Field(name = "country")
    private String country;

    @Field(name = "card_number")
    private String cardNumber;

    @Field(name = "vehicle")
    private String vehicle;

    @Field(name = "product_id")
    private String productId;

    @Field(name = "product_name")
    private String productName;

    @Field(name = "total_net_invoiced")
    private Double totalNetInvoiced;

    @Field(name = "total_vat_invoiced")
    private Double totalVatInvoiced;

    @Field(name = "currency")
    private String currency;
}
