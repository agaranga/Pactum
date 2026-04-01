FROM mcr.microsoft.com/dotnet/sdk:10.0-preview AS build
WORKDIR /src
COPY Pactum.Showcase/Pactum.Showcase.csproj Pactum.Showcase/
RUN dotnet restore Pactum.Showcase/Pactum.Showcase.csproj
COPY Pactum.Showcase/ Pactum.Showcase/
RUN dotnet publish Pactum.Showcase/Pactum.Showcase.csproj -c Release -o /app

FROM mcr.microsoft.com/dotnet/aspnet:10.0-preview
WORKDIR /app
COPY --from=build /app .
ENV ASPNETCORE_URLS=http://+:${PORT:-8080}
ENTRYPOINT ["dotnet", "Pactum.Showcase.dll"]
