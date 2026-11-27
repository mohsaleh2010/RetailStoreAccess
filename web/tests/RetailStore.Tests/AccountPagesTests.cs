using System.Net;
using System.Text.RegularExpressions;

namespace RetailStore.Tests;

/// <summary>The static pages of AccountLayout on MudBlazor (docs/ui-inventory.md section 16): the login is still a plain
/// form post with the field names the model binds, and an unknown address shows the not-found page.</summary>
public class AccountPagesTests
{
    private static async Task<string> TextOf(HttpResponseMessage r) => WebUtility.HtmlDecode(await r.Content.ReadAsStringAsync());

    [Fact]
    public async Task TheLoginFormPostsTheSameFieldNames()
    {
        using var host = new ApiHost();
        var client = host.CreateClient();
        var page = await TextOf(await client.GetAsync("/account/login"));
        Assert.Contains("name=\"Input.Username\"", page);
        Assert.Contains("name=\"Input.Password\"", page);
        Assert.Contains("mud-input-root-outlined", page);
        Assert.DoesNotContain("retailstore.css", page);

        string token = Regex.Match(page, "name=\"__RequestVerificationToken\" value=\"([^\"]+)\"").Groups[1].Value;
        var post = await client.PostAsync("/account/login", new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["_handler"] = "login",
            ["__RequestVerificationToken"] = token,
            ["Input.Username"] = "no-such-user",
            ["Input.Password"] = "x",
        }));
        var answer = await TextOf(post);
        Assert.Equal(HttpStatusCode.OK, post.StatusCode);
        Assert.Contains("mud-alert", answer);
        Assert.Contains("value=\"no-such-user\"", answer);
    }

    [Fact]
    public async Task AnUnknownAddressShowsTheNotFoundPage()
    {
        using var host = new ApiHost();
        var r = await host.CreateClient().GetAsync("/no-such-page");
        Assert.Equal(HttpStatusCode.NotFound, r.StatusCode);
        Assert.Contains("الصفحة غير موجودة", await TextOf(r));
    }
}
