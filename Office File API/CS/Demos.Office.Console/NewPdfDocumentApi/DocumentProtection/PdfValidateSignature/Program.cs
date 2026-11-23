using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Signed.pdf");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Validation_Result.txt";
    const bool ValidateCrlOcsp = true;
    const bool ValidateLtv = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = ValidateSignatures(documentStream, ValidateCrlOcsp, ValidateLtv);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream ValidateSignatures(Stream documentStream, bool validateClrOcsp, bool validateLtv) {
        var outputStream = new MemoryStream();
        var streamWriter = new StreamWriter(outputStream);

        using var doc = new PdfDocument(documentStream);

        if (doc.TryGetSignatures(out var signatureInfo)) {
            streamWriter.WriteLine($"Found {signatureInfo.Count} signatures");
            streamWriter.WriteLine();

            var signatureValidator = new PdfSignatureValidator(documentStream);

            for (var i = 0; i < signatureInfo.Count; i++) {
                ValidateSignature(signatureValidator, signatureInfo[i], i, validateClrOcsp, streamWriter);
            }

            if (validateLtv && signatureInfo.Count > 0)
                ValidateLtvInformation(signatureValidator, streamWriter);
        } else {
            streamWriter.WriteLine($"Signatures not found");
            streamWriter.WriteLine();
        }
        streamWriter.Flush();
        outputStream.Position = 0;
        return outputStream;
    }

    static void ValidateSignature(PdfSignatureValidator signatureValidator, SignatureInfo signature, int signatureIndex, bool validateClrOcsp, StreamWriter streamWriter) {
        streamWriter.WriteLine($"Signature #{signatureIndex}");
        streamWriter.WriteLine($"===");
        streamWriter.WriteLine($"Field Name : {signature.Field.Name}");
        streamWriter.WriteLine($"Signer name : {signature.SignerName}");

        var pkcs7 = signatureValidator.GetPkcs7Signature(signature.Field.Name);
        streamWriter.WriteLine($"Valid : {pkcs7.IsSignatureVerified}");

        var certificate = pkcs7.SignatureCertificate;
        streamWriter.WriteLine($"Certificate Issuer : {certificate.IssuerName.Name}");
        streamWriter.WriteLine($"Certificate Valid : {certificate.Verify()}");

        var timestampDate = pkcs7.TimeStampDate;
        streamWriter.WriteLine($"Has Timestamp : {timestampDate.HasValue}");
        if (timestampDate.HasValue) {
            streamWriter.WriteLine($"Timestamp : {timestampDate}");
            streamWriter.WriteLine($"Timestamp Valid : {pkcs7.IsTimeStampVerified}");
            streamWriter.WriteLine($"Document Timestamp : {pkcs7.IsDocumentTimeStamp}");
        }

        if (validateClrOcsp)
            ValidateCrlOcspInformation(pkcs7, streamWriter);

        streamWriter.WriteLine();
    }

    static void ValidateCrlOcspInformation(Pkcs7Signature pkcs7, StreamWriter streamWriter) {
        streamWriter.WriteLine();
        streamWriter.WriteLine($"CRL/OCSP Information");
        streamWriter.WriteLine($"CRL Embedded : {pkcs7.CrlList != null}");
        streamWriter.WriteLine($"OCSP Embedded : {pkcs7.OcspList != null}");

        bool tryFetchCrlOnline = false;
        bool tryFetchOcspOnline = false;

        var crlOcspResult = pkcs7.CheckCertificateRevocation(DateTime.UtcNow, tryFetchCrlOnline, tryFetchOcspOnline);

        streamWriter.WriteLine($"Certificate Revoked : {crlOcspResult.IsCrlRevoked}");
        if (crlOcspResult.IsCrlRevoked)
            streamWriter.WriteLine($"CRL Found Online : {crlOcspResult.IsCrlFoundOnline}");
        streamWriter.WriteLine($"OCSP Response Status : {crlOcspResult.OcspRevocationStatus}");
        if (crlOcspResult.OcspRevocationStatus != OcspRevocationStatus.None)
            streamWriter.WriteLine($"OCSP found online : {crlOcspResult.IsOcspFoundOnline}");
    }

    static void ValidateLtvInformation(PdfSignatureValidator signatureValidator, StreamWriter streamWriter) {
        streamWriter.WriteLine($"===");
        streamWriter.WriteLine($"LTV Information");
        streamWriter.WriteLine($"===");
        streamWriter.WriteLine();

        bool verifyEntireCertificateChain = true;
        bool tryFetchCrlOnline = false;
        bool tryFetchOcspOnline = false;

        var ltvResult = signatureValidator.VerifyLtv(verifyEntireCertificateChain, tryFetchCrlOnline, tryFetchOcspOnline);
        foreach (var entry in ltvResult) {
            streamWriter.WriteLine($"Signature Name : {entry.SignatureFieldName}");
            streamWriter.WriteLine($"Certificates In Chain : {entry.CertificateRevocationStatuses.Count}");

            foreach (var revocation in entry.CertificateRevocationStatuses) {
                streamWriter.WriteLine($"Certificate {revocation.Key.Subject}");
                streamWriter.WriteLine($"Certificate revoked : {revocation.Value.IsCrlRevoked}");
                if (revocation.Value.IsCrlRevoked)
                    streamWriter.WriteLine($"CRL found online : {revocation.Value.IsCrlFoundOnline}");
                streamWriter.WriteLine($"OCSP Response Status : {revocation.Value.OcspRevocationStatus}");
                if (revocation.Value.OcspRevocationStatus != OcspRevocationStatus.None)
                    streamWriter.WriteLine($"OCSP found online : {revocation.Value.IsOcspFoundOnline}");
                streamWriter.WriteLine();
            }
        }

        streamWriter.WriteLine();
    }
}
