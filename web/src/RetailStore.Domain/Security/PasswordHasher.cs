using System.Security.Cryptography;
using System.Text;

namespace RetailStore.Domain.Security;

/// <summary>The password digests of modSecurity (tools/security_reference.py): the password itself is never stored.
/// Employees.PasswordHash = lower-case hex of SHA-256 applied AppData.PasswordIterations times to UTF-8(salt + password),
/// each round hashing the 32 bytes of the previous one; Employees.PasswordSalt = 32 upper-case hex characters.</summary>
public static class PasswordHasher
{
    public static string Hash(string password, string salt)
    {
        byte[] data = Encoding.UTF8.GetBytes(salt + password);
        for (int i = 0; i < AppData.PasswordIterations; i++)
        {
            data = SHA256.HashData(data);
        }
        return Convert.ToHexStringLower(data);
    }

    public static string Sha256Hex(string text) => Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(text)));

    /// <summary>modSecurity.NewSalt: a random UUID without dashes, upper case.</summary>
    public static string NewSalt() => Guid.NewGuid().ToString("N").ToUpperInvariant();

    /// <summary>Compares two digests in constant time.</summary>
    public static bool Matches(string password, string salt, string? storedHash) =>
        storedHash is not null &&
        CryptographicOperations.FixedTimeEquals(Encoding.ASCII.GetBytes(Hash(password, salt)),
                                                Encoding.ASCII.GetBytes(storedHash));
}
