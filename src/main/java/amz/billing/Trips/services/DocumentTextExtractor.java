package amz.billing.Trips.services;

import java.io.InputStream;

public interface DocumentTextExtractor {

    String extractText(InputStream documentStream);
}