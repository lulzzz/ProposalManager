/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {
        status: {
            failed: "failed",
            succeeded: "succeeded"
        },
        sourceTypes: {
            all: 0,
            point: 1,
            table: 2,
            chart: 3
        },
        destinationTypes: {
            all: 0,
            point: 1,
            tableCell: 2,
            tableImage: 3,
            chart: 4
        },
        sortTypes: {
            df: "default",
            name: "name",
            status: "status",
            type: "type",
            value: "value"
        },
        sortOrder: {
            asc: "asc",
            desc: "desc"
        },
        intelliSenseResults: 6
    };

    app.guid = function () {
        var guid = "";
        for (var i = 1; i <= 32; i++) {
            var n = Math.floor(Math.random() * 16.0).toString(16);
            guid += n;
            if ((i == 8) || (i == 12) || (i == 16) || (i == 20))
                guid += "-";
        }
        return guid;
    };

    app.search = {
        splitKeyword: function (options) {
            var _k = options.keyword.toLocaleLowerCase(), _ka = [];
            $.each(_k.split(" "), function (x, y) {
                if (y != "") {
                    _ka.push(y);
                }
            });
            return _ka;
        },
        keyword: function (options) {
            var _ka = options.keyword, _ki = [], _a = [], _ra = [];
            for (var i = 0; i < _ka.length; i++) {
                _ki.push(String.fromCharCode(97 + i));
            }
            var _ks = _ki.join("");
            for (var i = _ka.length; i > 0; i--) {
                for (var m = 0; m < i; m++) {
                    _a.push(_ks.substr(m, _ka.length - i + 1));
                }
            }
            for (var i = 0; i < _a.length; i++) {
                var _ta = _a[i].split(""), _tl = _a[i].length, _tat = [];
                for (var m = 0; m < _ta.length; m++) {
                    _tat.push(_ka[_ta[m].charCodeAt() - 97]);
                }
                var _tatr = _tat.join(" ");
                _ra.push({ value: _tatr, length: _tl });
            }
            return _ra;
        },
        weight: function (options) {
            var _k = options.keyword, _n = $.trim(options.source.Name).toLocaleLowerCase(), _v = $.trim(options.source.Value).toLocaleLowerCase(), _f = 0;
            for (var i = 0; i < _k.length; i++) {
                if (options.source.SourceType == app.sourceTypes.point) {
                    if (_n.indexOf(_k[i]) > -1 || _v.indexOf(_k[i]) > -1) {
                        _f++;
                    }
                }
                else {
                    if (_n.indexOf(_k[i]) > -1) {
                        _f++;
                    }
                }
            }
            return _f >= _k.length ? 1 : 0;
        },
        autoComplete: function (options) {
            var _k = app.search.splitKeyword({ keyword: options.keyword }), _d = options.data, _r = [];
            $.each(_d, function (i, d) {
                var _w = app.search.weight({ keyword: _k, source: d });
                if (_w > 0) {
                    d.weight = _w;
                    _r.push(d);
                }
            });
            if (_r.length > 0) {
                options.result.find("li").remove();
                $.each(_r, function (i, d) {
                    if (i < app.intelliSenseResults) {
                        $('<li>' + d.Name + '</li>').appendTo(options.result);
                    }
                });
                options.result.show();
            }
            else {
                options.result.hide();
            }
            options.target.data("keyword", options.keyword);
        },
        move: function (options) {
            if (!options.result.is(":hidden")) {
                var _i = options.result.find("li.active").index(), _l = options.result.find("li").length, _m = false;
                if (options.down) {
                    if (_i < _l - 1) {
                        _i++;
                        _m = true;
                    }
                }
                else {
                    if (_i == -1) {
                        _i = _l;
                    }
                    if (_i > 0) {
                        _i--;
                        _m = true;
                    }
                }
                if (_m) {
                    options.result.find("li.active").removeClass("active");
                    options.result.find("li").eq(_i).addClass("active");
                    options.target.val(options.result.find("li").eq(_i).text());
                }
                else {
                    options.result.find("li.active").removeClass("active");
                    options.target.val(options.target.data("keyword"));
                }
            }
        }
    };

    app.string = function (str) {
        if (typeof (str) == "undefined" || str == null) {
            return "";
        }
        return $.trim(str);
    };

    app.format = function (n) {
        return n > 9 ? n : ("0" + n);
    },

    app.date = function (text) {
        var _v = new Date(text), _d = _v.getDate(), _m = _v.getMonth() + 1, _y = _v.getFullYear(), _h = _v.getHours(), _mm = _v.getMinutes(), _a = _h < 12 ? " AM" : " PM";
        return app.format(_m) + "/" + app.format(_d) + "/" + _y + " " + (_h < 12 ? (_h == 0 ? "12" : app.format(_h)) : (_h == 12 ? _h : _h - 12)) + ":" + app.format(_mm) + "" + _a + " PST";
    }

    app.sourceTitle = function (sourceType) {
        return sourceType == app.sourceTypes.point ? "point" : (sourceType == app.sourceTypes.table ? "table" : "chart")
    }

    app.mappings = {
        horizontalAlignment: {
            "General": "Left",
            "Left": "Left",
            "Center": "Centered",
            "CenterAcrossSelection": "Centered",
            "Right": "Right",
            "Justify": "Justified",
            "Fill": "Left",
            "Distributed": "Left"
        },
        verticalAlignment: {
            "Top": "Top",
            "Center": "Center",
            "Bottom": "Bottom",
            "Justify": "Bottom",
            "Distributed": "Bottom"
        },
        underline: {
            "None": "None",
            "Single": "Single",
            "SingleAccountant": "Single",
            "Double": "Double",
            "DoubleAccountant": "Double"
        },
        sideIndex: {
            "EdgeTop": "Top",
            "EdgeBottom": "Bottom",
            "EdgeLeft": "Left",
            "EdgeRight": "Right",
            "InsideVertical": "InsideVertical",
            "InsideHorizontal": "InsideHorizontal",
            "DiagonalDown": "",
            "DiagonalUp": ""
        },
        border: {
            "NoneThin": "None",
            "ContinuousHairline": "Dotted",
            "DotThin": "Dashed",
            "DashDotDotThin": "Dot2Dashed",
            "DashDotThin": "DotDashed",
            "DashThin": "DashedSmall",
            "ContinuousThin": "Single",
            "SlantDashDotMedium": "DashDotStroked",
            "ContinuousThick": "ThreeDEmboss",
            "DoubleThick": "Double",
            "DashDotDotMedium": "None",
            "DashDotMedium": "None",
            "DashMedium": "None",
            "ContinuousMedium": "None"
        }
    };

    app.formats = function (excelFormats) {
        var _formats = [];
        for (var i = 0; i < excelFormats.length; i++) {
            var _wordFormat = {}, _excelFormat = excelFormats[i];
            _wordFormat.columnWidth = Math.ceil(_excelFormat.columnWidth);
            _wordFormat.preferredHeight = Math.ceil(_excelFormat.rowHeight);
            _wordFormat.horizontalAlignment = app.mappings.horizontalAlignment[_excelFormat.horizontalAlignment];
            _wordFormat.verticalAlignment = app.mappings.verticalAlignment[_excelFormat.verticalAlignment];
            _wordFormat.shadingColor = _excelFormat.fill.color;
            _wordFormat.font = {
                bold: _excelFormat.font.bold,
                color: _excelFormat.font.color,
                italic: _excelFormat.font.italic,
                name: _excelFormat.font.name,
                size: _excelFormat.font.size,
                underline: app.mappings.underline[_excelFormat.font.underline]
            };
            _wordFormat.border = {};
            $.each(_excelFormat.borders, function (i, d) {
                if (d.sideIndex == "EdgeTop") {
                    _wordFormat.border.top = { color: d.color, type: app.mappings.border[d.style + d.weight] };
                }
                else if (d.sideIndex == "EdgeBottom") {
                    _wordFormat.border.bottom = { color: d.color, type: app.mappings.border[d.style + d.weight] };
                }
                else if (d.sideIndex == "EdgeLeft") {
                    _wordFormat.border.left = { color: d.color, type: app.mappings.border[d.style + d.weight] };
                }
                else if (d.sideIndex == "EdgeRight") {
                    _wordFormat.border.right = { color: d.color, type: app.mappings.border[d.style + d.weight] };
                }
                //else if (d.sideIndex == "InsideVertical") {
                //    _wordFormat.border.insideVertical = { color: d.color, type: app.mappings.border[d.style + d.weight] };
                //}
                //else if (d.sideIndex == "InsideHorizontal") {
                //    _wordFormat.border.insideHorizontal = { color: d.color, type: app.mappings.border[d.style + d.weight] };
                //}
            });
            _formats.push(_wordFormat);
        }
        return _formats;
    }

    return app;
})();