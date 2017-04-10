var IE = {
    READYSTATE_UNINITIALIZED: 0,
    READYSTATE_LOADING: 1,
    READYSTATE_LOADED: 2,
    READYSTATE_INTERACTIVE: 3,
    READYSTATE_COMPLETE: 4,

    _new: function(obj) {
        var ie = {
            _obj: obj,

            open: function(url) {
                if (this._obj != null) {
                    this._obj.Visible = true;
                    this._obj.Navigate(url);
                    while (this._obj.Busy || this._obj.ReadyState != IE.READYSTATE_COMPLETE) {
                        WScript.Sleep(100);
                    }
                }
            },

            quit: function() {
                if (this._obj != null) {
                    this._obj.Stop();
                    this._obj.Quit();
                }
                this._obj = null;
            },

            findText: function(text) {
                if (this._obj != null) {
                    return this._findText(this._obj.Document, text, "");
                } else {
                    return "";
                }
            },

            _findText: function(document, text, frameIds) {
                var paths = "";
                for (var element in document.all) {
                    if (element.innerText == text || element.value == text) {
                        var path = _keyValue("frame", frameIds);
                        path = _join(path, _keyValue("id", element.id), ",");
                        path = _join(path, _keyValue("name", element.name), ",");
                        path = _join(path, _keyValue("tagName", element.tagName), ",");
                        path = _join(path, _keyValue("innerText", element.innerText), ",");
                        path = _join(path, _keyValue("value", element.value), ",");
                        path = _join(path, _keyValue("text", element.text), ",");
                        paths = paths + path + "\n";
                    }
                }
                for (var i = 0; i < document.frames.length; i++){
                    // TODO: document.frames[i]
                    paths = paths + _findText(document.frames[i].document, text, _join(frameIds, i.toString(), "|"));
                }
                return paths;
            },

            _keyValue: function(key, value) {
                if (value == null || value == "") {
                    return "";
                } else {
                    return key + ":" + value;
                }
            },

            _join: function(str1, str2, separator) {
                if (str1 != null && str1 != "" && str2 != null && str2 != "") {
                    return str1 + separator + str2;
                } else {
                    return str1 + str2;
                }
            }
        };

        return ie;
    },

    create: function() {
        return this._new(new ActiveXObject("InternetExplorer.Application"));
    },

    find: function(regex) {
        var re = new RegExp(regex, "i");
        var shell = new ActiveXObject("Shell.Application");
        var windows = shell.Windows();
        if (windows != null) {
            for (var i = windows.Count - 1; i >= 0; i--) {
                var window = windows.Item(i);
                if (window != null) {
                    if (window.FullName.search(/iexplore\.exe/i)
                            && (window.LocationURL.search(re) != -1
                            || window.LocationName.search(re) != -1)) {
                        return this._new(window);
                    }
                }
            }
        }
        return null;
    }
};
