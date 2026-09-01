Attribute VB_Name = "调车计划登录"
'    _exportAsImage: function(b) {
'        var d = {
'            "sessionID": FR.SessionMgr.getSessionID(),
'            "width": this.width || 0,
'            "height": this.height || 0,
'            "index": b,
'            "__time": new Date().getTime()
'        };
'        $.extend(d, this.idInfo);
'        var a = FR.servletURL + "?op=chart&cmd=export_image";
'        for (var c in d) {
'            a += ("&" + c + "=" + d[c])
'        }
'        Window.Location = a
'    }


'curl 'http://10.190.48.4:8080/webroot/decision/login' \
'  -H 'Accept: application/json, text/javascript, */*; q=0.01' \
'  -H 'Accept-Language: zh-CN,zh;q=0.9' \
'  -H 'Connection: keep-alive' \
'  -H 'Content-Type: application/json' \
'  -H 'Origin: http://10.190.48.4:8080' \
'  -H 'Referer: http://10.190.48.4:8080/webroot/decision/login' \
'  -H 'User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36' \
'  -H 'X-Requested-With: XMLHttpRequest' \
'  -H 'transEncryptLevel: 1' \
'  --data-raw '{"username":"曾祥昊","password":"hsTly+cImEA2Tr7r6kwaM9dbXV2OM85wW8clMqYff72yhhxDE5DPFOldhXNQ9xdTEMbfA+WwjwwOS9rD5jQfpy4WbBrN6E/Y+t0D48EVhOzlGCqko0CRfQonAyNpn0/KfJfdDEFFnFtZlv5g75M3gi+sQMa96uxOSxUKCegTTCK+csGNbZVhe68L4/6Cuc0ZeH20NiFCsnEgS5KL5YwIkLMrWHypFVKoFPSVQADSUhGN1Xkz/XIvglhm3YiUWOu123b6FDXJlHhHe9I+/WsAR/9O4/usDDDeulmkVUhxU4JrYKcuZzEJDhkmhx92Jr0Heg7aKpKemOeTzd3azmBa6w==","validity":-1,"sliderToken":"","origin":"","encrypted":true}' \
'  --compressed \
'  --insecure
