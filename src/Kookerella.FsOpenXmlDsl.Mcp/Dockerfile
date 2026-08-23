# Build from the repo root as context, since this project references its sibling
# Kookerella.FsOpenXmlDsl library via a plain ProjectReference:
#   docker build -f src/Kookerella.FsOpenXmlDsl.Mcp/Dockerfile -t fsopenxmldsl-mcp .
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src
COPY src/Kookerella.FsOpenXmlDsl/Kookerella.FsOpenXmlDsl.fsproj src/Kookerella.FsOpenXmlDsl/
COPY src/Kookerella.FsOpenXmlDsl.Mcp/Kookerella.FsOpenXmlDsl.Mcp.fsproj src/Kookerella.FsOpenXmlDsl.Mcp/
RUN dotnet restore src/Kookerella.FsOpenXmlDsl.Mcp/Kookerella.FsOpenXmlDsl.Mcp.fsproj
COPY src/Kookerella.FsOpenXmlDsl/ src/Kookerella.FsOpenXmlDsl/
COPY src/Kookerella.FsOpenXmlDsl.Mcp/ src/Kookerella.FsOpenXmlDsl.Mcp/
RUN dotnet publish src/Kookerella.FsOpenXmlDsl.Mcp/Kookerella.FsOpenXmlDsl.Mcp.fsproj -c Release -o /app --no-restore

FROM mcr.microsoft.com/dotnet/runtime:10.0 AS runtime
WORKDIR /app
COPY --from=build /app .
# MCP over stdio - no exposed port. The client launches this and talks JSON-RPC over
# stdin/stdout (see Program.fs's own comment on why logging is routed to stderr instead).
ENTRYPOINT ["dotnet", "Kookerella.FsOpenXmlDsl.Mcp.dll"]
