FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src
COPY Pactum.Showcase/Pactum.Showcase.csproj Pactum.Showcase/
RUN sed -i 's/net10.0/net9.0/' Pactum.Showcase/Pactum.Showcase.csproj
RUN dotnet restore Pactum.Showcase/Pactum.Showcase.csproj
COPY Pactum.Showcase/ Pactum.Showcase/
RUN sed -i 's/net10.0/net9.0/' Pactum.Showcase/Pactum.Showcase.csproj
RUN dotnet publish Pactum.Showcase/Pactum.Showcase.csproj -c Release -o /app

FROM mcr.microsoft.com/dotnet/aspnet:9.0
WORKDIR /app
COPY --from=build /app .
ENV ASPNETCORE_URLS=http://+:${PORT:-8080}
ENTRYPOINT ["dotnet", "Pactum.Showcase.dll"]
