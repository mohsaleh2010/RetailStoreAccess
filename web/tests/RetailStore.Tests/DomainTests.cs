using RetailStore.Domain;
using RetailStore.Domain.Common;
using RetailStore.Domain.Localization;
using RetailStore.Domain.Security;

namespace RetailStore.Tests;

public class PasswordTests
{
    [Fact]
    public void Digests_are_the_ones_of_Access_and_Python()
    {
        foreach (var (password, salt, hash) in Vectors.Passwords)
        {
            Assert.Equal(hash, PasswordHasher.Hash(password, salt));
            Assert.True(PasswordHasher.Matches(password, salt, hash));
            Assert.False(PasswordHasher.Matches(password + "x", salt, hash));
        }
        foreach (var (text, sha) in Vectors.Sha256)
        {
            Assert.Equal(sha, PasswordHasher.Sha256Hex(text));
        }
    }

    [Fact]
    public void Salt_is_32_upper_case_hex_characters()
    {
        string salt = PasswordHasher.NewSalt();
        Assert.Matches("^[0-9A-F]{32}$", salt);
        Assert.NotEqual(salt, PasswordHasher.NewSalt());
    }

    [Fact]
    public void Password_rules()
    {
        Assert.NotEqual("", PasswordRules.Problem("12345", "ali"));
        Assert.NotEqual("", PasswordRules.Problem("Ahmed1", "ahmed1"));
        Assert.NotEqual("", PasswordRules.Problem(" secret1", "ali"));
        Assert.Equal("", PasswordRules.Problem("secret1", "ali"));
        Assert.Equal(6, AppData.MinPasswordLength);
    }
}

public class CountryTests
{
    [Fact]
    public void Tax_numbers_and_mobiles_of_each_country()
    {
        Assert.Equal("", CountryRules.TaxNumberProblem(null, "SA"));
        Assert.Equal("", CountryRules.TaxNumberProblem("300000000000003", "SA"));
        Assert.NotEqual("", CountryRules.TaxNumberProblem("300000000000004", "SA"));
        Assert.NotEqual("", CountryRules.TaxNumberProblem("123456789", "SA"));
        Assert.Equal("", CountryRules.TaxNumberProblem("123456789", "EG"));
        Assert.True(CountryRules.MobileFits("0501234567", "SA"));
        Assert.True(CountryRules.MobileFits("+966501234567", "SA"));
        Assert.False(CountryRules.MobileFits("01012345678", "SA"));
        Assert.True(CountryRules.MobileFits("01012345678", "EG"));
        Assert.Equal("SA", CountryRules.Normalize(null));
        Assert.True(CountryRules.EmailFits("a@b.co"));
        Assert.False(CountryRules.EmailFits("a@b"));
        Assert.False(CountryRules.EmailFits("@b.co"));
    }

    [Fact]
    public void Money_rounds_half_away_from_zero()
    {
        Assert.Equal(2.35m, MoneyMath.RoundMoney(2.345m));
        Assert.Equal(-2.35m, MoneyMath.RoundMoney(-2.345m));
    }
}

public class TranslatorTests
{
    [Fact]
    public void English_is_what_modLang_Tr_gives()
    {
        foreach (var (arabic, english) in Vectors.Translations)
        {
            Assert.Equal(english, Translator.ToEnglish(arabic));
        }
    }

    [Fact]
    public void A_key_glued_to_other_letters_is_not_replaced()
    {
        Assert.Equal("xالمنتجاتy", Translator.ReplaceWord("xالمنتجاتy", "المنتج", "Product"));
        Assert.Equal("Product x", Translator.ReplaceWord("المنتج x", "المنتج", "Product"));
        Assert.Equal("abc", Translator.ToEnglish("abc"));
    }
}
