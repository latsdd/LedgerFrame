using System;

namespace LedgerFrame
{
    // Development stub: passthrough (no decryption).
    // Release builds embed an encrypted model and replace this file
    // with a real key holder as part of a private packaging step.
    static class KeyHolder
    {
        public static byte[] Decrypt(byte[] data)
        {
            return data;
        }
    }
}
