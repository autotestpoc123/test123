 // Known image-data failures only. Do not turn NAS I/O, missing codecs or
        // arbitrary runtime failures into a corrupt-photo result.
        private static bool IsInvalidImage(Exception ex) => ex is InvalidDataException
            || ex is COMException && (uint)ex.HResult is
                0x88982F07 or // WINCODEC_ERR_UNKNOWNIMAGEFORMAT
                0x88982F60 or // WINCODEC_ERR_BADIMAGE
                0x88982F61 or // WINCODEC_ERR_BADHEADER
                0x88982F62 or // WINCODEC_ERR_FRAMEMISSING
                0x88982F63;   // WINCODEC_ERR_BADMETADATAHEADER
