package service;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class DriveService {

    private final Drive service;

    public DriveService(Drive service) {
        this.service = service;
    }

    public File copy(String template, String newFileName) throws IOException {
        File copyMetadata = new File().setName(newFileName);
        return service.files().copy(template, copyMetadata).execute();
    }

    public void exportPDF(String presentationId, String fileName) throws IOException {
        OutputStream out = new FileOutputStream(fileName);
        service.files().export(presentationId, "application/pdf").executeMediaAndDownloadTo(out);
    }
}
