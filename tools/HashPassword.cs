// Simple tool to generate password hashes
// Run: dotnet script tools/HashPassword.cs <password>
// Or: dotnet run --project tools/HashPassword.csproj <password>

using Microsoft.AspNetCore.Identity;

var password = args.Length > 0 ? args[0] : "changeme";
var hasher = new PasswordHasher<object>();
var hash = hasher.HashPassword(new object(), password);
Console.WriteLine(hash);
