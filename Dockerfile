# syntax=docker/dockerfile:1.7

FROM node:22-bookworm-slim AS ui-build
WORKDIR /src/web
COPY web/package.json web/package-lock.json ./
RUN npm ci
COPY web/ ./
COPY Samples/ /src/Samples/
ARG APP_VERSION=dev
RUN MONOREPORT_VERSION="$APP_VERSION" npm run build

FROM mcr.microsoft.com/dotnet/sdk:10.0 AS api-build
WORKDIR /src
COPY global.json Directory.Build.props NuGet.Config ./
COPY MonoReport/MonoReport.csproj MonoReport/
COPY MonoReport.Api/MonoReport.Api.csproj MonoReport.Api/
RUN dotnet restore MonoReport.Api/MonoReport.Api.csproj
COPY MonoReport/ MonoReport/
COPY MonoReport.Api/ MonoReport.Api/
COPY --from=ui-build /src/web/dist /ui-dist
ARG APP_VERSION=dev
RUN dotnet publish MonoReport.Api/MonoReport.Api.csproj -c Release --no-restore -o /app/publish \
    -p:MonoReportSkipUiBuild=true -p:MonoReportUiDist=/ui-dist -p:MonoReportUiVersion="$APP_VERSION"

FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS runtime
ARG SOURCE_URL=https://github.com/montr/monoreport
ARG APP_VERSION=dev
LABEL org.opencontainers.image.source="$SOURCE_URL" \
    org.opencontainers.image.version="$APP_VERSION" \
    org.opencontainers.image.title="MonoReport"
WORKDIR /app
COPY --from=api-build --chown=$APP_UID:$APP_UID /app/publish ./
ENV ASPNETCORE_HTTP_PORTS=8080 \
    DOTNET_EnableDiagnostics=0
EXPOSE 8080
USER $APP_UID
ENTRYPOINT ["dotnet", "MonoReport.Api.dll"]
