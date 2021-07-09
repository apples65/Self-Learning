FROM mcr.microsoft.com/dotnet/sdk:3.1-alpine

# install IPA fonts
RUN apk add --no-cache curl fontconfig &&\
    curl -O https://moji.or.jp/wp-content/ipafont/IPAfont/IPAfont00303.zip && \
    mkdir -p /usr/share/fonts/IPAfont && \
    mkdir -p /temp && \
    unzip IPAfont00303.zip -d /temp && \
    cp /temp/IPAfont00303/ipag.ttf /usr/share/fonts/IPAfont/ && \
    cp /temp/IPAfont00303/ipagp.ttf /usr/share/fonts/IPAfont/ && \
    cp /temp/IPAfont00303/ipam.ttf /usr/share/fonts/IPAfont/ && \
    cp /temp/IPAfont00303/ipamp.ttf /usr/share/fonts/IPAfont/ && \
    rm IPAfont00303.zip
RUN rm -rf /temp && \
    fc-cache -fv 