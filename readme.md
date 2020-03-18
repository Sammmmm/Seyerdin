# Building

SSubTmr6.dll needs to be added as a reference

scivbx.ocx and vbaListView6.ocx do too, could use some steps here to do this without banging your head against a wall

# Development

There are several command line arguments for the client that are useful for local development

- \-iP \<ipaddress\> - change target ip
- \-port \<port\> - change target port
- \-serverid \<serverid\> - this will append the string <serverid> to relevant files such as the map cache and character caches, to prevent weirdness between servers.  It will load skilldata<serverid>.dat and classes<serverid>.ini from the cache, to load the relevant settings for a server.   
