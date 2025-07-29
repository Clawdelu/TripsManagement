package amz.billing.Trips.services;

import org.springframework.web.multipart.MultipartFile;

public interface IDkvService {

    byte[] processDetailedDkvRaport(MultipartFile dkvRaport);
}
