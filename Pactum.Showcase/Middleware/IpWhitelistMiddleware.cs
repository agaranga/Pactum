using System.Net;

namespace Pactum.Showcase.Middleware;

public class IpWhitelistMiddleware
{
    private readonly RequestDelegate _next;
    private readonly HashSet<string> _allowedIps;
    private readonly HashSet<string> _allowedSubnets;
    private readonly ILogger<IpWhitelistMiddleware> _logger;

    public IpWhitelistMiddleware(RequestDelegate next, IConfiguration config, ILogger<IpWhitelistMiddleware> logger)
    {
        _next = next;
        _logger = logger;

        var ips = config.GetSection("IpWhitelist:AllowedIPs").Get<string[]>() ?? [];
        _allowedIps = new HashSet<string>(ips, StringComparer.OrdinalIgnoreCase);

        var subnets = config.GetSection("IpWhitelist:AllowedSubnets").Get<string[]>() ?? [];
        _allowedSubnets = new HashSet<string>(subnets, StringComparer.OrdinalIgnoreCase);
    }

    public async Task InvokeAsync(HttpContext context)
    {
        var remoteIp = context.Connection.RemoteIpAddress;

        if (remoteIp == null)
        {
            context.Response.StatusCode = StatusCodes.Status403Forbidden;
            return;
        }

        // Always allow localhost
        if (IPAddress.IsLoopback(remoteIp))
        {
            await _next(context);
            return;
        }

        var ipString = remoteIp.MapToIPv4().ToString();

        if (_allowedIps.Contains(ipString))
        {
            await _next(context);
            return;
        }

        foreach (var subnet in _allowedSubnets)
        {
            if (IsInSubnet(remoteIp, subnet))
            {
                await _next(context);
                return;
            }
        }

        _logger.LogWarning("Blocked request from {IP}", ipString);
        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        await context.Response.WriteAsync("Access denied");
    }

    private static bool IsInSubnet(IPAddress address, string cidr)
    {
        var parts = cidr.Split('/');
        if (parts.Length != 2 || !IPAddress.TryParse(parts[0], out var subnet) || !int.TryParse(parts[1], out var bits))
            return false;

        var addressBytes = address.MapToIPv4().GetAddressBytes();
        var subnetBytes = subnet.MapToIPv4().GetAddressBytes();

        if (addressBytes.Length != subnetBytes.Length)
            return false;

        var mask = bits == 0 ? 0u : uint.MaxValue << (32 - bits);
        var addressInt = BitConverter.ToUInt32(addressBytes.Reverse().ToArray(), 0);
        var subnetInt = BitConverter.ToUInt32(subnetBytes.Reverse().ToArray(), 0);

        return (addressInt & mask) == (subnetInt & mask);
    }
}
