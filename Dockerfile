# syntax=docker/dockerfile:1
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src
COPY src/ExchangeEwsMcp/ExchangeEwsMcp.csproj src/ExchangeEwsMcp/
RUN dotnet restore src/ExchangeEwsMcp/ExchangeEwsMcp.csproj
COPY src/ src/
ARG VERSION=0.1.0
RUN dotnet publish src/ExchangeEwsMcp/ExchangeEwsMcp.csproj -c Release -o /app --no-restore -p:Version=${VERSION}

FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
WORKDIR /app
COPY --from=build /app .
ENV ASPNETCORE_HTTP_PORTS=8080 \
    Server__DataDirectory=/data
RUN mkdir -p /data && chown -R $APP_UID /data
VOLUME /data
USER $APP_UID
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s CMD ["dotnet", "ExchangeEwsMcp.dll", "--healthcheck"]
ENTRYPOINT ["dotnet", "ExchangeEwsMcp.dll"]
