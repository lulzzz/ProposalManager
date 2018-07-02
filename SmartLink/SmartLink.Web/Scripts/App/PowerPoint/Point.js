$(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            point.init();
        });
    };
});

var point = (function () {
    var point = {
        filePath: "",
        controls: {},
        selectDocument: null,
        points: [],
        sourcePointKeyword: "",
        pagerIndex: 0,
        pagerSize: 30,
        pagerCount: 0,
        totalPoints: 0,
        endpoints: {
            catalog: "/api/SourcePointCatalog?documentId=",
            token: "/api/GraphAccessToken",
            sharePointToken: "/api/SharePointAccessToken",
            graph: "https://graph.microsoft.com/v1.0",
            userInfo: "/api/userprofile"
        },
        api: {
            host: "",
            token: "",
            sharePointToken: ""
        }
    }, that = point;

    that.init = function () {
        that.filePath = Office.context.document.url.indexOf("E:") > -1 ? "https://cand3.sharepoint.com/Shared%20Documents/Presentation.pptx" : Office.context.document.url;
        that.controls = {
            body: $("body"),
            main: $(".main"),
            select: $(".n-select"),
            refresh: $(".n-refresh"),
            titleName: $("#lblSourcePointName"),
            sourcePointName: $("#txtSearchSourcePoint"),
            searchSourcePoint: $("#iSearchSourcePoint"),
            autoCompleteControl2: $("#autoCompleteWrap2"),
            list: $("#listPoints"),
            headerListPoints: $("#headerListPoints"),
            popupMain: $("#popupMain"),
            popupErrorOK: $("#btnErrorOK"),
            popupMessage: $("#popupMessage"),
            popupProcessing: $("#popupProcessing"),
            popupSuccessMessage: $("#lblSuccessMessage"),
            popupErrorMain: $("#popupErrorMain"),
            popupErrorTitle: $("#lblErrorTitle"),
            popupErrorMessage: $("#lblErrorMessage"),
            popupErrorRepair: $("#lblErrorRepair"),
            popupBrowseList: $("#browseList"),
            popupBrowseBack: $("#btnBrowseBack"),
            popupBrowseCancel: $("#btnBrowseCancel"),
            popupBrowseMessage: $("#txtBrowseMessage"),
            popupBrowseLoading: $("#popBrowseLoading"),
            /*Message Bar begin*/
            innerMessageBox: $("#innerMessageBox"),
            innerMessageIcon: $("#innerMessageIcon"),
            innerMessageText: $("#innerMessageText"),
            /*Message Bar end*/
            pager: $("#pager"),
            pagerTotal: $("#pagerTotal"),
            pagerPages: $("#pagerPages"),
            pagerCurrent: $("#pagerCurrent"),
            pagerPrev: $("#pagerPrev"),
            pagerNext: $("#pagerNext"),
            pagerValue: $("#pagerValue"),
            pagerGo: $("#pagerGo"),
            indexes: $("#indexes"),
            /* Footer begin */
            footer: $(".footer"),
            settings: $(".settings"),
            openSettings: $(".f-settings"),
            closeSettings: $(".s-settings"),
            userName: $(".s-username"),
            email: $(".s-email")
        };

        that.controls.body.click(function () {
            that.action.body();
        });
        that.controls.select.click(function () {
            that.browse.init();
        });
        that.controls.refresh.click(function () {
            if (that.selectDocument != null) {
                that.list();
            }
            else {
                that.popup.message({ success: false, title: "Please select file first." });
            }
        });
        that.controls.sourcePointName.focus(function () {
            that.action.dft(this, true);
        });
        that.controls.sourcePointName.blur(function () {
            that.action.dft(this, false);
        });
        that.controls.sourcePointName.keydown(function (e) {
            if (e.keyCode == 13) {
                if (that.controls.sourcePointName.val() != "") {
                    $(".search-tooltips").hide();
                    that.action.searchSourcePoint(true);
                }
            }
            else if (e.keyCode == 38 || e.keyCode == 40) {
                app.search.move({ result: that.controls.autoCompleteControl2, target: that.controls.sourcePointName, down: e.keyCode == 40 });
            }
        });
        that.controls.sourcePointName.bind("input", function (e) {
            that.action.autoComplete2();
        });
        that.controls.searchSourcePoint.click(function () {
            that.action.searchSourcePoint(!that.controls.sourcePointName.closest(".search").hasClass("searched"));
        });
        that.controls.popupErrorOK.click(function () {
            that.action.ok();
        });
        that.controls.popupBrowseList.on("click", "li", function () {
            that.browse.select($(this));
        });
        that.controls.popupBrowseCancel.click(function () {
            that.browse.popup.hide();
        });
        that.controls.popupBrowseBack.click(function () {
            that.browse.popup.back();
        });

        that.controls.innerMessageBox.on("click", ".close-Message", function () {
            that.popup.hide();
            return false;
        });

        that.controls.list.on("click", ".i-history", function () {
            that.action.history($(this).closest(".point-item"));
            return false;
        });
        that.controls.main.on("click", ".search-tooltips li", function () {
            $(this).parent().parent().find("input").val($(this).text());
            $(this).parent().hide();
            that.action.searchSourcePoint(true);
        });
        that.controls.main.on("mouseover", ".search-tooltips li", function () {
            $(this).parent().find("li.active").removeClass("active");
            $(this).addClass("active");
        });
        that.controls.main.on("mouseout", ".search-tooltips li", function () {
            $(this).removeClass("active");
        });
        that.controls.pagerPrev.on("click", function () {
            if (!$(this).hasClass("disabled")) {
                that.utility.pager.prev();
            }
        });
        that.controls.pagerNext.on("click", function () {
            if (!$(this).hasClass("disabled")) {
                that.utility.pager.next();
            }
        });
        that.controls.pagerValue.on("keydown", function (e) {
            if (e.keyCode == 13) {
                that.controls.pagerGo.click();
            }
        });
        that.controls.pagerGo.on("click", function () {
            var _v = $.trim(that.controls.pagerValue.val());
            if (isNaN(_v)) {
                that.popup.message({ success: false, title: "Only numbers are a valid input." });
            }
            else {
                _n = parseInt(_v);
                if (_n > 0 && _n <= that.pagerCount) {
                    that.utility.pager.init({ index: _n });
                }
                else {
                    that.popup.message({ success: false, title: "Invalid number." });
                }
            }
        });

        /* Footer begin */
        that.controls.settings.blur(function () {
            if (that.controls.footer.hasClass("footer-shorter")) {
                that.controls.footer.removeClass("footer-shorter");
                that.controls.settings.removeClass("show-settings");
            }
        });
        that.controls.openSettings.click(function () {
            that.controls.footer.addClass("footer-shorter");
            that.controls.settings.addClass("show-settings");
        });
        that.controls.closeSettings.click(function () {
            that.controls.footer.removeClass("footer-shorter");
            that.controls.settings.removeClass("show-settings");
        });
        /* Footer end */

        $(window).resize(function () {
            that.utility.height();
        });
        that.utility.height();
        that.action.dft(that.controls.sourcePointName, false);
        that.utility.pager.status({ length: 0 });
        
            // Get user Info
            that.userInfo(function (result) {
                if (result.status == app.status.succeeded) {
                    that.controls.userName[0].innerText = result.data.Username;
                    that.controls.email[0].innerText = result.data.Email;
                }
            });
        
    };

    that.userInfo = function (callback) {
        that.service.userInfo(function (result) {
            callback(result);
        });
    };

    that.list = function () {
        that.popup.processing(true);
        that.service.catalog({ documentId: that.selectDocument.Id }, function (result) {
            if (result.status == app.status.succeeded) {
                that.popup.processing(false);
                that.controls.titleName.html("Source Points in " + that.selectDocument.Name);
                that.controls.titleName.prop("title", "Source Points in " + that.selectDocument.Name);
                that.points = result.data && result.data.SourcePoints && result.data.SourcePoints.length > 0 ? result.data.SourcePoints : [];
                if (that.points.length > 0) {
                    that.utility.pager.init({ index: 1 });
                }
                else {
                    that.utility.pager.status({ length: 0 });
                }
            }
            else {
                that.popup.message({ success: false, title: "Load source points failed." });
            }
        });
    };

    that.utility = {
        format: function (n) {
            return n > 9 ? n : ("0" + n);
        },
        date: function (str) {
            var _v = new Date(str), _d = _v.getDate(), _m = _v.getMonth() + 1, _y = _v.getFullYear(), _h = _v.getHours(), _mm = _v.getMinutes(), _a = _h < 12 ? " AM" : " PM";
            return that.utility.format(_m) + "/" + that.utility.format(_d) + "/" + _y + " " + (_h < 12 ? (_h == 0 ? "12" : that.utility.format(_h)) : (_h == 12 ? _h : _h - 12)) + ":" + that.utility.format(_mm) + "" + _a + " PST";
        },
        position: function (p) {
            if (p != null && p != undefined) {
                var _i = p.lastIndexOf("!"), _s = p.substr(0, _i).replace(new RegExp(/('')/g), '\''), _c = p.substr(_i + 1, p.length);
                if (_s.indexOf("'") == 0) {
                    _s = _s.substr(1, _s.length);
                }
                if (_s.lastIndexOf("'") == _s.length - 1) {
                    _s = _s.substr(0, _s.length - 1);
                }
                return { sheet: _s, cell: _c };
            }
            else {
                return { sheet: "", cell: "" };
            }
        },
        add: function (options) {
            that.points.push(options);
        },
        fileName: function (path) {
            path = decodeURI(path);
            return path.lastIndexOf("/") > -1 ? path.substr(path.lastIndexOf("/") + 1) : (path.lastIndexOf("\\") > -1 ? path.substr(path.lastIndexOf("\\") + 1) : path);
        },
        pager: {
            init: function (options) {
                that.controls.pagerValue.val("");
                that.controls.indexes.html("");
                that.pagerIndex = options.index ? options.index : 1;
                that.ui.list();
            },
            prev: function () {
                that.controls.pagerValue.val("");
                that.utility.pager.updatePager();
                that.pagerIndex--;
                that.ui.list();
            },
            next: function () {
                that.controls.pagerValue.val("");
                that.utility.pager.updatePager();
                that.pagerIndex++;
                that.ui.list();
            },
            status: function (options) {
                that.totalPoints = options.length;
                that.pagerCount = Math.ceil(options.length / that.pagerSize);
                that.controls.pagerTotal.html(options.length);
                that.controls.pagerPages.html(that.pagerCount);
                that.controls.pagerCurrent.html(that.pagerIndex);
                that.utility.pager.updatePager();
                that.pagerIndex == 1 || that.pagerCount == 0 ? that.controls.pagerPrev.addClass("disabled") : that.controls.pagerPrev.removeClass("disabled");
                that.pagerIndex == that.pagerCount || that.pagerCount == 0 ? that.controls.pagerNext.addClass("disabled") : that.controls.pagerNext.removeClass("disabled");
                if (that.totalPoints <= that.pagerSize) {
                    that.controls.pagerPrev.removeClass("disabled").addClass("disabled");
                    that.controls.pagerNext.removeClass("disabled").addClass("disabled");
                }
                else {
                    if (that.pagerIndex == 1) {
                        that.controls.pagerPrev.removeClass("disabled").addClass("disabled");
                    }
                    if (that.pagerIndex == that.pagerCount) {
                        that.controls.pagerNext.removeClass("disabled").addClass("disabled");
                    }
                }
            },
            updatePager: function () {
                var _start = ((that.pagerIndex - 1) * that.pagerSize + 1);
                var _left = that.totalPoints - _start;
                var _end = _left < 0 ? 0 : (_left <= that.pagerSize ? (_start + _left) : (_start + that.pagerSize - 1));
                if (_end > 0) {
                    that.controls.indexes.html(_start + "-" + _end);
                }
                else {
                    that.controls.indexes.html("");
                }
            }
        },
        height: function () {
            /*if (that.controls.main.hasClass("manage")) {
                var _h = that.controls.main.outerHeight();
                var _h1 = $("#pager").outerHeight();
                that.controls.list.css("maxHeight", (_h - 206 - 70 - _h1) + "px");
            }*/
        },
        path: function () {
            var _a = that.filePath.split("//")[1], _b = _a.split("/"), _p = [_b[0]];
            _b.shift();
            _b.pop();
            for (var i = 1; i <= _b.length; i++) {
                var _c = [];
                for (var n = 0; n < i; n++) {
                    _c.push(_b[n]);
                }
                _p.push(_c.join("/"));
            }
            return _p;
        },
        publishHistory: function (options) {
            var _ph = [], _tmd = options.data && options.data.length > 0 ? $.extend([], options.data) : [], _td = _tmd.reverse();
            $.each(_td, function (m, n) {
                var __c = _td[m].Value ? _td[m].Value : "",
                    __p = _td[m > 0 ? m - 1 : m].Value ? _td[m > 0 ? m - 1 : m].Value : "";
                if (m == 0 || __c != __p) {
                    _ph.push(n);
                }
            });
            return _ph.reverse().slice(0, 5);
        }
    };

    that.action = {
        body: function () {
            $(".search-tooltips").hide();
        },
        dft: function (elem, on) {
            var _k = $.trim($(elem).val()), _kd = $(elem).data("default");
            if (on) {
                if (_k == _kd) {
                    $(elem).val("");
                }
                $(elem).removeClass("input-default");
            }
            else {
                if (_k == "" || _k == _kd) {
                    $(elem).val(_kd).addClass("input-default");
                }
            }
        },
        history: function (o) {
            o.hasClass("item-more") ? o.removeClass("item-more") : o.addClass("item-more");
        },
        ok: function () {
            that.controls.popupMain.removeClass("active message process confirm");
        },
        autoComplete2: function () {
            var _e = $.trim(that.controls.sourcePointName.val()), _d = that.points;
            if (_e != "") {
                app.search.autoComplete({ keyword: _e, data: _d, result: that.controls.autoCompleteControl2, target: that.controls.sourcePointName });
            }
            else {
                that.controls.autoCompleteControl2.hide();
            }
        },
        searchSourcePoint: function (s) {
            that.sourcePointKeyword = $.trim(that.controls.sourcePointName.val()) == that.controls.sourcePointName.data("default") ? "" : $.trim(that.controls.sourcePointName.val());
            if (s && that.sourcePointKeyword != "") {
                that.controls.sourcePointName.closest(".search").addClass("searched");
                that.controls.searchSourcePoint.removeClass("ms-Icon--Search").addClass("ms-Icon--Cancel");
            }
            else {
                that.sourcePointKeyword = "";
                that.controls.sourcePointName.val("");
                that.action.dft(that.controls.sourcePointName, false);
                that.controls.sourcePointName.closest(".search").removeClass("searched");
                that.controls.searchSourcePoint.removeClass("ms-Icon--Cancel").addClass("ms-Icon--Search");
            }
            that.utility.pager.init({});
        }
    };

    that.browse = {
        path: [],
        init: function () {
            that.api.token = "";
            that.browse.path = [];
            that.browse.popup.dft();
            that.browse.popup.show();
            that.browse.popup.processing(true);
            that.browse.token();
        },
        token: function () {
            that.service.token({ endpoint: that.endpoints.token }, function (result) {
                if (result.status == app.status.succeeded) {
                    that.api.token = result.data;
                    that.api.host = that.utility.path()[0].toLowerCase();
                    that.service.token({ endpoint: that.endpoints.sharePointToken }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.api.sharePointToken = result.data;
                            that.browse.siteCollection();
                        }
                        else {
                            that.document.error({ title: "Get sharepoint access token failed." });
                        }
                    });
                }
                else {
                    that.browse.popup.message("Get graph access token failed.");
                }
            });
        },
        siteCollection: function (options) {
            if (typeof (options) == "undefined") {
                options = {
                    path: that.utility.path().reverse(),
                    index: 0,
                    values: [],
                    webUrls: []
                };
            }
            if (options.index < options.path.length) {
                that.service.siteCollection({ path: options.path[options.index] }, function (result) {
                    if (result.status == app.status.succeeded) {
                        if (typeof (result.data.siteCollection) != "undefined") {
                            options.values.push(result.data.id);
                            options.webUrls.push(result.data.webUrl);
                        }
                        options.index++;
                        that.browse.siteCollection(options);
                    }
                    else {
                        if (result.error.status == 401) {
                            that.browse.popup.message("Access denied.");
                        }
                        else {
                            options.index++;
                            that.browse.siteCollection(options);
                        }
                    }
                });
            }
            else {
                if (options.values.length > 0) {
                    that.browse.sites({ siteId: options.values.shift(), siteUrl: options.webUrls.shift() });
                }
                else {
                    that.browse.popup.message("Get site collection ID failed.");
                }
            }
        },
        sites: function (options) {
            that.service.sites(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _s = [];
                    $.each(result.data.value, function (i, d) {
                        _s.push({ id: d.id, name: d.name, type: "site", siteUrl: d.webUrl });
                    });
                    that.browse.libraries({ siteId: options.siteId, siteUrl: options.siteUrl, sites: _s });
                }
                else {
                    that.browse.popup.message("Get sites failed.");
                }
            });
        },
        libraries: function (options) {
            that.service.libraries(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _l = options.sites ? options.sites : [];
                    $.each(result.data.value, function (i, d) {
                        _l.push({ id: d.id, name: decodeURI(d.name), type: "library", siteId: options.siteId, siteUrl: options.siteUrl, url: d.webUrl });
                    });
                    that.browse.display({ data: _l });
                }
                else {
                    that.browse.popup.message("Get libraries failed.");
                }
            });
        },
        items: function (options) {
            if (options.inFolder) {
                that.service.itemsInFolder(options, function (result) {
                    if (result.status == app.status.succeeded) {
                        var _fd = [], _fi = [];
                        $.each(result.data.value, function (i, d) {
                            var _u = d.webUrl, _n = d.name, _nu = decodeURI(_n);
                            if (d.folder) {
                                _fd.push({ id: d.id, name: _nu, type: "folder", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                            }
                            else if (d.file) {
                                if (_n.toUpperCase().indexOf(".XLSX") > 0) {
                                    _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                }
                            }
                        });
                        _fi.sort(function (_a, _b) {
                            return (_a.name.toUpperCase() > _b.name.toUpperCase()) ? 1 : (_a.name.toUpperCase() < _b.name.toUpperCase()) ? -1 : 0;
                        });
                        that.browse.display({ data: _fd.concat(_fi) });
                    }
                    else {
                        that.browse.popup.message("Get files failed.");
                    }
                });
            }
            else {
                that.service.items(options, function (result) {
                    if (result.status == app.status.succeeded) {
                        var _fd = [], _fi = [];
                        $.each(result.data.value, function (i, d) {
                            var _u = d.webUrl, _n = d.name, _nu = decodeURI(_n);
                            if (d.folder) {
                                _fd.push({ id: d.id, name: _nu, type: "folder", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                            }
                            else if (d.file) {
                                if (_n.toUpperCase().indexOf(".XLSX") > 0) {
                                    _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                }
                            }
                        });
                        _fi.sort(function (_a, _b) {
                            return (_a.name.toUpperCase() > _b.name.toUpperCase()) ? 1 : (_a.name.toUpperCase() < _b.name.toUpperCase()) ? -1 : 0;
                        });
                        that.browse.display({ data: _fd.concat(_fi) });
                    }
                    else {
                        that.browse.popup.message("Get files failed.");
                    }
                });
            }
        },
        file: function (options) {
            that.service.item(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _d = "";
                    $.each(result.data.value, function (i, d) {
                        if (decodeURI(options.url).toUpperCase() == decodeURI(d.EncodedAbsUrl).toUpperCase() && d.OData__dlc_DocId) {
                            _d = d.OData__dlc_DocId;
                            return false;
                        }
                    });
                    if (_d != "") {
                        that.browse.popup.hide();
                        that.selectDocument = { Id: _d, Name: options.fileName };
                        that.list();
                    }
                    else {
                        that.browse.popup.message("Get file Document ID failed.");
                    }
                }
                else {
                    that.browse.popup.message("Get file Document ID failed.");
                }
            });
        },
        display: function (options) {
            that.controls.popupBrowseList.html("");
            $.each(options.data, function (i, d) {
                var _h = "";
                if (d.type == "site") {
                    _h = '<li class="i-site" data-id="' + d.id + '" data-type="site" data-siteurl="' + d.siteUrl + '">' + d.name + '</li>';
                }
                else if (d.type == "library") {
                    _h = '<li class="i-library" data-id="' + d.id + '" data-site="' + d.siteId + '" data-url="' + d.url + '" data-type="library" data-siteurl="' + d.siteUrl + '" data-listname="' + d.name + '">' + d.name + '</li>';
                }
                else if (d.type == "folder") {
                    _h = '<li class="i-folder" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="folder" data-siteurl="' + d.siteUrl + '" data-listname="' + d.listName + '">' + d.name + '</li>';
                }
                else if (d.type == "file") {
                    _h = '<li class="i-file" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="file" data-siteurl="' + d.siteUrl + '" data-listname="' + d.listName + '">' + d.name + '</li>';
                }
                that.controls.popupBrowseList.append(_h);
            });
            if (options.data.length == 0) {
                that.controls.popupBrowseList.html("No items found.");
            }
            that.browse.popup.processing(false);
        },
        select: function (elem) {
            var _t = $(elem).data("type");
            if (_t == "site") {
                that.browse.path.push({ type: "site", id: $(elem).data("id"), siteUrl: $(elem).data("siteurl") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.sites({ siteId: $(elem).data("id"), siteUrl: $(elem).data("siteurl") });
            }
            else if (_t == "library") {
                that.browse.path.push({ type: "library", id: $(elem).data("id"), site: $(elem).data("site"), url: $(elem).data("url"), siteUrl: $(elem).data("siteurl"), listName: $(elem).data("listname") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.items({ inFolder: false, siteId: $(elem).data("site"), siteUrl: $(elem).data("siteurl"), listId: $(elem).data("id"), listName: $(elem).data("listname") });
            }
            else if (_t == "folder") {
                that.browse.path.push({ type: "folder", id: $(elem).data("id"), site: $(elem).data("site"), siteUrl: $(elem).data("siteurl"), list: $(elem).data("list"), url: $(elem).data("url"), listName: $(elem).data("listname") });
                that.browse.popup.nav();
                that.browse.popup.processing(true);
                that.browse.items({ inFolder: true, siteId: $(elem).data("site"), siteUrl: $(elem).data("siteurl"), listId: $(elem).data("list"), listName: $(elem).data("listname"), itemId: $(elem).data("id") });
            }
            else {
                that.browse.popup.processing(true);
                that.browse.file({ siteUrl: $(elem).data("siteurl"), listName: $(elem).data("listname"), name: $(elem).text(), url: that.browse.path[that.browse.path.length - 1].url + "/" + encodeURI($(elem).text()), fileName: $.trim($(elem).text()) });
            }
        },
        popup: {
            dft: function () {
                that.controls.popupBrowseList.html("");
                that.controls.popupBrowseBack.hide();
                that.controls.popupBrowseMessage.html("").hide();
                that.controls.popupBrowseLoading.hide();
            },
            show: function () {
                that.controls.popupMain.removeClass("message process confirm").addClass("active browse");
            },
            hide: function () {
                that.controls.popupMain.removeClass("active message process confirm browse");
            },
            processing: function (show) {
                if (show) {
                    that.controls.popupBrowseLoading.show();
                }
                else {
                    that.controls.popupBrowseLoading.hide();
                }
            },
            message: function (txt) {
                that.controls.popupBrowseMessage.html(txt).show();
                that.browse.popup.processing(false);
            },
            back: function () {
                that.browse.path.pop();
                if (that.browse.path.length > 0) {
                    var _ip = that.browse.path[that.browse.path.length - 1];
                    if (_ip.type == "site") {
                        that.browse.popup.processing(true);
                        that.browse.sites({ siteId: _ip.id, siteUrl: _ip.siteUrl });
                    }
                    else if (_ip.type == "library") {
                        that.browse.popup.processing(true);
                        that.browse.items({ inFolder: false, siteId: _ip.site, listId: _ip.id, siteUrl: _ip.siteUrl, listName: _ip.listName });
                    }
                    else if (_ip.type == "folder") {
                        that.browse.popup.processing(true);
                        that.browse.items({ inFolder: true, siteId: _ip.site, listId: _ip.list, itemId: _ip.id, siteUrl: _ip.siteUrl, listName: _ip.listName });
                    }
                }
                else {
                    that.browse.popup.processing(true);
                    that.browse.siteCollection();
                }
                that.browse.popup.nav();
            },
            nav: function () {
                that.browse.path.length > 0 ? that.controls.popupBrowseBack.show() : that.controls.popupBrowseBack.hide();
            }
        }
    };

    that.popup = {
        message: function (options, callback) {
            $(".popups .bg").removeAttr("style");
            if (options.success) {
                that.controls.popupMessage.removeClass("error").addClass("success");
                // that.controls.popupSuccessMessage.html(options.title);
                that.controls.innerMessageBox.removeClass("active ms-MessageBar--error").addClass("active ms-MessageBar--success");
                that.controls.innerMessageIcon.removeClass("ms-Icon--ErrorBadge").addClass("ms-Icon--Completed");
                that.controls.innerMessageText.html(options.title);
                $(".popups .bg").hide();
            }
            else {
                if (options.values) {
                    that.controls.popupMessage.removeClass("success").addClass("error");
                    that.controls.popupErrorTitle.html(options.title ? options.title : "");
                    var _s = "error-single";

                    _s = "error-list";
                    var _h = "";
                    $.each(options.values, function (i, d) {
                        _h += "<li>";
                        $.each(d, function (m, n) {
                            _h += "<span>" + n + "</span>";
                        });
                        _h += "</li>";
                    });
                    that.controls.popupErrorMessage.html(_h);
                    if (options.repair) {
                        _s = "error-repair";
                        that.controls.popupErrorRepair.html(options.repair);
                    }

                    that.controls.popupErrorMain.removeClass("error-single error-list").addClass(_s);
                }
                else {
                    that.controls.popupMessage.removeClass("error").addClass("success");
                    // that.controls.popupSuccessMessage.html(options.title);
                    that.controls.innerMessageBox.removeClass("active ms-MessageBar--success").addClass("active ms-MessageBar--error");
                    that.controls.innerMessageIcon.removeClass("ms-Icon--Completed").addClass("ms-Icon--ErrorBadge");
                    that.controls.innerMessageText.html(options.title);
                    $(".popups .bg").hide();
                }
            }
            if (options.canClose) {
                that.controls.innerMessageBox.addClass("canclose");
            }
            else {
                that.controls.innerMessageBox.removeClass("canclose");
            }
            that.controls.popupMain.removeClass("process confirm browse").addClass("active message");
            if (options.success) {
                callback();
            }
            else {
                if (options.values) {
                    that.controls.popupErrorOK.unbind("click").click(function () {
                        that.action.ok();
                        if (callback) {
                            callback();
                        }
                    });
                }
                else {
                    if (callback) {
                        callback();
                    }
                    else {
                        that.popup.hide(3000);
                    }
                }
            }
        },
        processing: function (show) {
            if (!show) {
                that.controls.popupMain.removeClass("active process");
            }
            else {
                that.controls.popupMain.removeClass("message confirm browse").addClass("active process");
            }
        },
        browse: function (show) {
            if (!show) {
                that.controls.popupMain.removeClass("active browse");
            }
            else {
                that.controls.popupMain.removeClass("message process confirm").addClass("active browse");
            }
        },
        hide: function (millisecond) {
            if (millisecond) {
                setTimeout(function () {
                    $(".popups .bg").removeAttr("style");
                    that.controls.popupMain.removeClass("active message");
                    that.controls.innerMessageBox.removeClass("active");
                }, millisecond);
            } else {
                $(".popups .bg").removeAttr("style");
                that.controls.popupMain.removeClass("active message");
                that.controls.innerMessageBox.removeClass("active");
            }
        },
        back: function (millisecond) {
            if (millisecond) {
                setTimeout(function () {
                    that.controls.popupMain.removeClass("active message");
                    that.controls.main.removeClass("manage add").addClass("manage");
                }, millisecond);
            }
            else {
                that.controls.popupMain.removeClass("active message");
                that.controls.main.removeClass("manage add").addClass("manage");
            }
        }
    };

    that.service = {
        common: function (options, callback) {
            $.ajax({
                url: options.url,
                type: options.type,
                cache: false,
                data: options.data ? options.data : "",
                dataType: options.dataType,
                headers: options.headers ? options.headers : "",
                success: function (data) {
                    callback({ status: app.status.succeeded, data: data });
                },
                error: function (error) {
                    if (error.status == 410) {
                        that.popup.message({ success: false, title: "The current login gets expired and needs re-authenticate. You will be redirected to the login page by click OK." }, function () {
                            window.location = "../../Home/Index";
                        });
                    }
                    else {
                        callback({ status: app.status.failed, error: error });
                    }
                }
            });
        },
        catalog: function (options, callback) {
            that.service.common({ url: that.endpoints.catalog + options.documentId, type: "GET", dataType: "json" }, callback);
        },
        token: function (options, callback) {
            that.service.common({ url: options.endpoint, type: "GET", dataType: "json" }, callback);
        },
        siteCollection: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + that.api.host + ":/" + options.path, type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        sites: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/sites", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        libraries: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/drives", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        items: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/drives/" + options.listId + "/root/children", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        itemsInFolder: function (options, callback) {
            that.service.common({ url: that.endpoints.graph + "/sites/" + options.siteId + "/drives/" + options.listId + "/items/" + options.itemId + "/children", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        item: function (options, callback) {
            that.service.common({ url: options.siteUrl + "/_api/web/lists/getbytitle('" + options.listName + "')/items?$select=FileLeafRef,EncodedAbsUrl,OData__dlc_DocId&$filter=FileLeafRef eq '" + options.fileName + "'", type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.sharePointToken } }, callback);
        },
        userInfo: function (callback) {
            that.service.common({ url: that.endpoints.userInfo, type: "GET" }, callback);
        }
    };

    that.ui = {
        dft: function () {
            var _f = $.trim(that.controls.file.val()), _fd = that.controls.file.data("default"), _k = $.trim(that.controls.keyword.val()), _kd = that.controls.keyword.data("default");
            if (_f == "" || _f == _fd) {
                that.controls.file.val(_fd);
            }
            if (_k == "" || _k == _kd) {
                that.controls.keyword.val(_kd).addClass("input-default");
                that.controls.search.removeClass("ms-Icon--Search ms-Icon--Cancel").addClass("ms-Icon--Search");
            }
        },
        list: function (options) {
            var _dt = $.extend([], that.points), _d = [], _ss = [];
            if (that.sourcePointKeyword != "") {
                var _sk = app.search.splitKeyword({ keyword: that.sourcePointKeyword });
                if (_sk.length > 26) {
                    that.popup.message({ success: false, title: "Only support less then 26 keywords." });
                }
                else {
                    $.each(_dt, function (i, d) {
                        if (app.search.weight({ keyword: _sk, source: d }) > 0) {
                            _d.push(d);
                        }
                    });
                }
            }
            else {
                _d = _dt;
            }
            that.utility.pager.status({ length: _d.length });
            _d.sort(function (_a, _b) {
                return (app.string(_a.Name).toUpperCase() > app.string(_b.Name).toUpperCase()) ? 1 : (app.string(_a.Name).toUpperCase() < app.string(_b.Name).toUpperCase()) ? -1 : 0;
            });
            that.controls.list.find(".point-item").remove();
            that.ui.item({ index: 0, data: _d, selected: _ss });
        },
        item: function (options, callback) {
            if (options.index < options.data.length) {
                var _item = options.data[options.index];
                if (options.index >= that.pagerSize * (that.pagerIndex - 1) && options.index < that.pagerSize * that.pagerIndex) {
                    var _pn = that.utility.position(_item.NamePosition),
                        _pv = (_item.PublishedHistories && _item.PublishedHistories.length > 0 ? (_item.PublishedHistories[0].Value ? _item.PublishedHistories[0].Value : "") : ""),
                        _pht = that.utility.publishHistory({ data: _item.PublishedHistories });
                    var _h = '<li class="point-item" data-id="' + _item.Id + '" data-range="' + _item.RangeId + '" data-position="' + _item.Position + '" data-namerange="' + _item.NameRangeId + '" data-nameposition="' + _item.NamePosition + '">';
                    _h += '<div class="point-item-line">';
                    _h += '<div class="i2"><span class="s-name" title="' + _item.Name + '">' + _item.Name + '</span>';
                    _h += '<div class="sp-file-pos">';
                    _h += '<span title="' + (_pn.sheet ? _pn.sheet : "") + ':[' + (_pn.cell ? _pn.cell : "") + ']">' + (_pn.sheet ? _pn.sheet : "") + ':[' + (_pn.cell ? _pn.cell : "") + ']</span>';
                    _h += '</div>';
                    _h += '</div>';
                    _h += '<div class="i3" title="' + _pv + '">' + _pv + '</div>';
                    _h += '<div class="i5">';
                    _h += '<div class="i-menu"><a href="javascript:"><span title="Action"><i class="ms-Icon ms-Icon--More"></i></span><span class="quick-menu"><span class="i-history" title="History"><i class="ms-Icon ms-Icon--ChevronRight"></i><i>History</i></span></span></a></div>';
                    _h += '</div>';
                    _h += '<div class="clear"></div>';
                    _h += '</div>';

                    _h += '<div class="item-history"><ul class="history-list">';
                    _h += '<li class="history-header"><div class="h1"><span>Name</span></div><div class="h2"><span>Value</span></div><div class="h3"><span>Date</span></div></li>';
                    $.each(_pht, function (m, n) {
                        _h += '<li class="history-item"><div class="h1" title="' + n.PublishedUser + '">' + n.PublishedUser + '</div><div class="h2" title="' + (n.Value ? n.Value : "") + '">' + (n.Value ? n.Value : "") + '</div><div class="h3" title="' + that.utility.date(n.PublishedDate) + '">' + that.utility.date(n.PublishedDate) + '</div></li>';
                    });
                    _h += '</ul>';
                    _h += '<div class="clear"></div>';
                    _h += '</div>';
                    _h += '</li>';
                    that.controls.list.append(_h);
                }
                options.index++;
                that.ui.item(options);
            }
            else {
                that.controls.main.get(0).scrollTop = 0;
            }
        }
    };

    return that;
})();