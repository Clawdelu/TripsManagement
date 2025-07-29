package amz.billing.Trips.enums;

import com.fasterxml.jackson.annotation.JsonValue;

public enum Status {
    COMPLETED("COMPLETED"),
    CANCELLED("CANCELLED"),
    DETENTION("DETENTION");

    private final String value;

    Status(String value) {
        this.value = value;
    }
    @JsonValue
    public String getValue() {
        return value;
    }
}
