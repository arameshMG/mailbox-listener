FROM mcr.microsoft.com/powershell:7.4-ubuntu-22.04

RUN pwsh -NoLogo -NonInteractive -Command \
    "Install-Module -Name ExchangeOnlineManagement -Force -Scope AllUsers -Repository PSGallery"

WORKDIR /app

COPY listener.ps1 /app/listener.ps1
COPY entrypoint.sh /app/entrypoint.sh

RUN chmod +x /app/entrypoint.sh

EXPOSE 5050

ENTRYPOINT ["/app/entrypoint.sh"]

