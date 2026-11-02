namespace RetailStore.Domain.Security;

/// <summary>modSecurity.PasswordProblem: "" when the password is acceptable, otherwise the reason (Arabic, through Tr).</summary>
public static class PasswordRules
{
    public static string Problem(string password, string username)
    {
        if (password.Length < AppData.MinPasswordLength)
        {
            return "كلمة المرور يجب ألا تقل عن " + AppData.MinPasswordLength + " أحرف.";
        }
        if (string.Equals(password.Trim(), username.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            return "كلمة المرور لا يجوز أن تكون نفس اسم المستخدم.";
        }
        if (password != password.Trim())
        {
            return "كلمة المرور لا تبدأ ولا تنتهي بمسافة.";
        }
        return "";
    }
}
