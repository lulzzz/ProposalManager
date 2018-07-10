$(function () {
	Office.initialize = function (reason)
	{
		var isInIframe = function ()
		{
			try 
			{
				return window.self !== window.top;
			}
			catch (e)
			{
				return true;
			}
		};

        $(document).ready(function () {
			if (isInIframe())
			{
				microsoftTeams.initialize();

				microsoftTeams.authentication.authenticate({
					url: '/auth',
					width: 600,
					height: 535,
					successCallback: function (result)
					{
						point.init(result.idToken);
						$("#dvLogin").hide();
						$("#excel-addin").show();

					},
					failureCallback: function (err)
					{
						console.log(err);
					}
				});
			}
			else
			{
				var authenticationContext = new AuthenticationContext(config);

				// Check For & Handle Redirect From AAD After Login
				if (authenticationContext.isCallback(window.location.hash))
				{
					authenticationContext.handleWindowCallback();
				}
				else 
				{
					var user = authenticationContext.getCachedUser();
					if (user && window.parent === window && !window.opener)
					{

						authenticationContext.acquireToken(config.clientId,
							function (errorDesc, token, error)
							{
								if (error)
								{
									authenticationContext.acquireTokenRedirect(config.clientId, null, null);
								}
								else
								{
									point.init(token);
									$("#dvLogin").hide();
									$("#excel-addin").show();
								}
							});
					}
					else
					{
						authenticationContext.login();
					}
				}
			}
		});
	}
});

var point = (function () {
    var point = {
        filePath: "",
        documentId: "",
        controls: {},
        points: [],
        filteredPoints: [],
        endpoints: {
            add: "/api/SourcePoint",
            edit: "/api/SourcePoint",
            list: "/api/SourcePointCatalog?",
            del: "/api/SourcePoint?id=",
            publish: "/api/PublishSourcePoints",
            associated: "/api/DestinationPoint?sourcePointId=",
            deleteSelected: "/api/DeleteSelectedSourcePoint",
            token: "/api/GraphAccessToken",
            sharePointToken: "/api/SharePointAccessToken",
            graph: "https://graph.microsoft.com/v1.0",
            checkCloneStatus: "/api/CloneCheckFile",
            cloneFiles: "/api/CloneFiles",
            userInfo: "/api/userprofile"
        },
        model: null,
        bulk: false,
        sourcePointKeyword: "",
        pagerIndex: 0,
        pagerSize: 30,
        pagerCount: 0,
        totalPoints: 0,
        isBulk: function () { return that.controls.main.hasClass("bulk") },
        isPoint: function () { return that.controls.main.hasClass("single") },
        isTable: function () { return that.controls.main.hasClass("table") },
        isChart: function () { return that.controls.main.hasClass("chart") },
        api: {
            host: "",
            token: "",
            sharePointToken: ""
        }
    }, that = point;

	that.init = function (accessToken)
	{
		that.token = accessToken;
		that.filePath = Office.context.document.url;
        that.controls = {
            body: $("body"),
            main: $(".main"),
            back: $(".n-back"),
            add: $(".n-add"),
            publish: $(".n-publish"),
            publishAll: $(".n-publishall"),
            refresh: $(".n-refresh"),
            del: $(".n-delete"),
            bulk: $(".n-bulk"),
            cancel: $("#btnCancel"),
            save: $("#btnSave"),
            name: $("#txtName"),
            position: $("#txtLocation"),
            select: $("#btnLocation"),
            selectName: $("#btnSelectName"),
            list: $("#listPoints"),
            documentIdError: $("#lblDocumentIDError"),
            documentIdReload: $("#btnDocumentIDReload"),
            headerListPoints: $("#headerListPoints"),
            sourcePointName: $("#txtSearchSourcePoint"),
            searchSourcePoint: $("#iSearchSourcePoint"),
            autoCompleteControl2: $("#autoCompleteWrap2"),
            listAssociated: $("#listAssociated"),
            moreMenu: $(".n-more"),
            moreMenuBox: $(".more-menu"),
            resort: $(".resort"),
            /* Clone begin */
            clone: $(".s-cloning"),
            sourceFolder: $("#txtSourceFolder"),
            sourceFolderTrigger: $("#btnOpenSourceFolder"),
            destinationFolder: $("#txtDestinationFolder"),
            destinationFolderTrigger: $("#btnOpenDestinationFolder"),
            sourceFolderName: $("#sourceFolderName"),
            destFolderName: $("#destFolderName"),
            cloneBtn: $("#btnClone"),
            cloneNext: $("#btnCloneNext"),
            cloneCancel: $("#btnCloneCancel"),
            cloneDone: $("#btnCloneDone"),
            listWillClone: $("#listWillClone"),
            listWillNotClone: $("#listWillNotClone"),
            listCloneError: $("#listCloneError"),
            listDoneSuccess: $("#listDoneSuccess"),
            listDoneFail: $("#listDoneFail"),
            linkCloned: $("#linkCloned"),
            cloneResult: $(".clone-point .clone-point-fourth"),
            /* Clone end */
            /* New sub folder begin */
            newFolder: $("#btnNewFolder"),
            newFolderName: $("#txtNewFolderName"),
            newFolderCreate: $("#btnCreateNewFolder"),
            newFolderMessage: $("#browseNewFolderMessage"),
            /* New sub folder end */
            /* Source table and chart being */
            sourceTypeNavMana: $(".point-types-mana li"),
            sourceTypeNav: $(".point-types-add li"),
            tableName: $("#txtTableTitle"),
            tableLocation: $("#txtTableLocation"),
            tableLocationSelect: $("#btnTableLocation"),
            chartName: $("#txtChartName"),
            chartLocation: $("#txtChartLocation"),
            chartLocationSelect: $("#btnChartLocation"),
            chartList: $("#chartList"),
            moveUp: $("#btnMoveUp"),
            moveDown: $("#btnMoveDown"),
            /* Source table and chart end */
            popupMain: $("#popupMain"),
            popupErrorOK: $("#btnErrorOK"),
            popupMessage: $("#popupMessage"),
            popupProcessing: $("#popupProcessing"),
            popupSuccessMessage: $("#lblSuccessMessage"),
            popupErrorMain: $("#popupErrorMain"),
            popupErrorTitle: $("#lblErrorTitle"),
            popupErrorMessage: $("#lblErrorMessage"),
            popupErrorRepair: $("#lblErrorRepair"),
            popupConfirm: $("#popupConfirm"),
            popupConfirmTitle: $("#lblConfirmTitle"),
            popupConfirmYes: $("#btnYes"),
            popupConfirmNo: $("#btnNo"),
            popupBrowse: $("#popupBrowse"),
            popupBrowseList: $("#browseList"),
            popupBrowseBack: $("#btnBrowseBack"),
            popupBrowseCancel: $("#btnBrowseCancel"),
            popupBrowseMessage: $("#txtBrowseMessage"),
            popupBrowseLoading: $("#popBrowseLoading"),
            popupBrowseOKWrap: $("#wrapBrowseOK"),
            popupBrowseOK: $("#btnBrowseOK"),
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
        that.controls.add.click(function () {
            that.action.add();
        });
        that.controls.back.click(function () {
            that.action.back();
        });
        that.controls.publish.click(function () {
            that.action.publish();
        });
        that.controls.publishAll.click(function () {
            that.action.publishAll();
            that.controls.moreMenuBox.hide();
        });
        that.controls.refresh.click(function () {
            that.action.refresh();
        });
        that.controls.del.click(function () {
            that.action.deleteSelected();
        });
        that.controls.bulk.click(function () {
            that.action.bulk();
        });
        that.controls.selectName.click(function () {
            that.action.select({ input: that.controls.name });
        });
        that.controls.select.click(function () {
            that.action.select({ input: that.controls.position });
        });
        that.controls.save.click(function () {
            that.action.save();
        });
        that.controls.cancel.click(function () {
            that.action.back();
        });
        that.controls.resort.click(function () {
            that.action.resort();
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
        that.controls.documentIdReload.click(function () {
            window.location.reload();
        });

        that.controls.popupBrowseList.on("click", "li", function () {
            that.browse.select($(this));
        });
        that.controls.popupBrowseList.on("click", "li.i-folder i", function () {
            that.browse.checked($(this).parent());
            return false;
        });
        that.controls.popupBrowseOK.on("click", function () {
            that.browse.selectFolder();
        });
        that.controls.popupBrowseCancel.click(function () {
            if (that.controls.popupBrowse.hasClass("newfolder")) {
                that.action.newFolder.cancel();
            }
            else {
                that.browse.popup.hide();
            }
        });
        that.controls.popupBrowseBack.click(function () {
            that.browse.popup.back();
        });

        that.controls.innerMessageBox.on("click", ".close-Message", function () {
            that.popup.hide();
            return false;
        });

        that.controls.moreMenu.mouseover(function () {
            that.controls.moreMenu.css({ "background-color": "#dadada" })
            that.controls.moreMenuBox.show();
        });
        that.controls.moreMenuBox.mouseleave(function () {
            that.controls.moreMenu.css({ "background-color": "#ffffff" })
            that.controls.moreMenuBox.hide();
        })

        /* Clone begin */
        that.controls.clone.click(function () {
            that.controls.footer.removeClass("footer-shorter");
            that.controls.settings.removeClass("show-settings");
            that.action.clone.init();
        });
        that.controls.cloneBtn.click(function () {
            if (!that.controls.cloneBtn.hasClass("disabled")) {
                that.action.clone.save();
            }
        });
        that.controls.cloneNext.click(function () {
            if (!that.controls.cloneNext.hasClass("disabled")) {
                that.action.clone.next();
            }
        });

        that.controls.cloneCancel.click(function () {
            // confirmation dialogue
            that.popup.confirm({ title: "Are you sure you want to cancel cloning?", message: "" }, function () {
                // back to list
                that.controls.popupMain.removeClass("message process confirm active");
                that.action.backToList();
            }, function () {
                that.controls.popupMain.removeClass("message process confirm active");
            });

        });
        that.controls.cloneDone.click(function () {
            // back to list
            that.action.backToList();
        });
        that.controls.sourceFolder.click(function () {
            $(this).closest(".clone-point-first").removeClass("source dest").addClass("source");
            that.browse.browseFile = false;
            //that.action.clone.source = null;
            //that.action.clone.destination = null;
            that.browse.init();
        });
        that.controls.sourceFolderTrigger.click(function () {
            $(this).closest(".clone-point-first").removeClass("source dest").addClass("source");
            that.browse.browseFile = false;
            that.controls.popupBrowse.removeClass("dpdeslibrary");
            //that.action.clone.source = null;
            //that.action.clone.destination = null;
            that.browse.init();
        });
        that.controls.destinationFolder.click(function () {
            $(this).closest(".clone-point-first").removeClass("source dest").addClass("dest");
            that.browse.browseFile = false;
            //that.action.clone.destination = null;
            that.browse.init();
        });
        that.controls.destinationFolderTrigger.click(function () {
            $(this).closest(".clone-point-first").removeClass("source dest").addClass("dest");
            that.browse.browseFile = false;
            that.controls.popupBrowse.addClass("dpdeslibrary");
            //that.action.clone.destination = null;
            that.browse.init();
        });
        /* Clone end */
        /* New folder begin */
        that.controls.newFolder.click(function () {
            that.action.newFolder.init();
        });
        that.controls.newFolderCreate.click(function () {
            that.action.newFolder.create();
        });
        /* New folder end */
        /* Source table and chart begin */

        that.controls.sourceTypeNavMana.click(function () {
            var _t = false;
            if ($(this).hasClass("is-selected")) {
                $(this).removeClass("is-selected");
                _t = true;
            } else {
                that.controls.sourceTypeNavMana.removeClass("is-selected");
                $(this).addClass("is-selected");
            }

            if ($(this).data("content") != "Points") {
                if (!_t) {
                    that.controls.headerListPoints.find(".point-header").removeClass("type-point type-table type-chart only").addClass("type-table only");
                    that.controls.headerListPoints.find(".i2 span")[0].innerText = "Name";
                }
                else {
                    that.controls.headerListPoints.find(".i2 span")[0].innerText = "Source Point";
                    that.controls.headerListPoints.find(".point-header").removeClass("type-point type-table type-chart only").addClass("type-table");
                }
            }
            else {
                that.controls.headerListPoints.find(".i2 span")[0].innerText = "Source Point";
                if (!_t) {
                    that.controls.headerListPoints.find(".point-header").removeClass("type-point type-table type-chart only").addClass("type-point only");
                }
                else {
                    that.controls.headerListPoints.find(".point-header").removeClass("type-point type-table type-chart only").addClass("type-point");
                }
            }
            $(".ckb-wrapper.all").find("input").prop("checked", false);
            $(".ckb-wrapper.all").removeClass("checked");
            that.utility.scrollTop();
            that.controls.list.find(".point-item").remove();
            that.controls.moveUp.removeClass("disabled").addClass("disabled");
            that.controls.moveDown.removeClass("disabled").addClass("disabled");
            that.utility.pager.init({ refresh: false });
        });

        that.controls.sourceTypeNav.click(function () {
            $(this).parent().find("li.is-selected").removeClass("is-selected");
            $(this).addClass("is-selected");
            if ($(this).index() == 0) {
                that.controls.main.removeClass("table chart").addClass("single");
                if (that.controls.main.hasClass("edit")) {
                    $(that.controls.back.find(".n-name-edit")[0]).text("Edit Source Point");
                }
                else {
                    $(that.controls.back.find(".n-name-add")[0]).text("Add Source Point");
                }
            }
            else if ($(this).index() == 1) {
                that.controls.main.removeClass("single table").addClass("chart");
                if (that.controls.main.hasClass("edit")) {
                    $(that.controls.back.find(".n-name-edit")[0]).text("Edit Source Chart");
                }
                else {
                    $(that.controls.back.find(".n-name-add")[0]).text("Add Source Chart");
                }
            }
            else if ($(this).index() == 2) {
                that.controls.main.removeClass("single chart").addClass("table");
                if (that.controls.main.hasClass("edit")) {
                    $(that.controls.back.find(".n-name-edit")[0]).text("Edit Source Table");
                }
                else {
                    $(that.controls.back.find(".n-name-add")[0]).text("Add Source Table");
                }
            }
        });
        that.controls.tableLocationSelect.click(function () {
            that.action.select({ input: that.controls.tableLocation });
        });
        that.controls.headerListPoints.find(".i2,.i3,.i4").click(function () {
            that.action.sort($(this));
        });
        that.controls.chartLocationSelect.click(function () {
            that.action.charts();
            return false;
        });
        that.controls.chartList.on("click", "li a", function () {
            that.action.selectChart($(this));
            return false;
        });
        /* Source table and chart end */
        that.controls.moveUp.click(function () {
            that.action.up();
        });
        that.controls.moveDown.click(function () {
            that.action.down();
        });
        that.controls.list.on("click", ".i-history", function () {
            that.action.history($(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".item-history", function () {
            that.action.history($(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".i-delete", function () {
            that.action.del($(this).closest(".point-item").data("id"), $(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".i-edit", function () {
            that.action.edit($(this).closest(".point-item").data("id"), $(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".point-item", function (e) {
            that.action.goto($(this).data("id"), $(this));
            if ($(this).closest(".point-item").hasClass("item-more")) {
                $(this).closest(".point-item").removeClass("item-more");
            }
        });
        that.controls.main.on("change", ".ckb-wrapper input", function () {
            if ($(this).get(0).checked) {
                $(this).closest(".ckb-wrapper").addClass("checked");
                if ($(this).closest(".all").length > 0) {
                    that.controls.list.find("input").prop("checked", true);
                    that.controls.list.find(".ckb-wrapper").addClass("checked");
                }
            }
            else {
                $(this).closest(".ckb-wrapper").removeClass("checked");
                if ($(this).closest(".all").length > 0) {
                    that.controls.list.find("input").prop("checked", false);
                    that.controls.list.find(".ckb-wrapper").removeClass("checked");
                }
            }
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
                    that.utility.pager.init({ index: _n, refresh: false });
                }
                else {
                    that.popup.message({ success: false, title: "Invalid number." });
                }
            }
        });
        that.errorFilter.init();


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
            that.controls.moreMenu.css({ "background-color": "#ffffff" })
            that.controls.moreMenuBox.hide();
        });
        that.utility.height();
        that.action.dft(that.controls.sourcePointName, false);

        // Retrieve the document ID via document URL
        that.document.init(function () {
            that.list({ refresh: true, index: 1 }, function (result) {
                if (result.status == app.status.failed) {
                    that.popup.message({ success: false, title: result.error.statusText });
                }
            });

            that.userInfo(function (result) {
                if (result.status == app.status.succeeded) {
                    that.controls.userName[0].innerText = result.data.Username;
                    that.controls.email[0].innerText = result.data.Email;
                }
            });
        });
    };

    that.userInfo = function (callback) {
        that.service.userInfo(function (result) {
            callback(result);
        });
    };

    that.list = function (options, callback) {
        that.popup.processing(true);
        that.service.list(function (result) {
            if (result.status == app.status.succeeded) {
                if (result.data) {
                    that.points = result.data.SourcePoints;
                    that.utility.pager.init({ refresh: options.refresh, index: options.index }, function () {
                        that.popup.processing(false);
                        callback({ status: result.status });
                    });
                }
                else {
                    that.utility.pager.status({ length: 0 });
                    that.popup.processing(false);
                    callback({ status: result.status });
                }
            }
            else {
                callback({ status: result.status, error: result.error });
            }
        });
    };

    that.default = function (callback) {
        if (that.model) {
            if (that.model.SourceType == app.sourceTypes.point) {
                that.controls.name.val(that.model.NamePosition);
                that.controls.position.val(that.model.Position);
                that.controls.tableName.val("");
                that.controls.tableLocation.val("");
                that.controls.chartName.val("");
                that.controls.chartLocation.val("");
            }
            else if (that.model.SourceType == app.sourceTypes.table) {
                that.controls.name.val("");
                that.controls.position.val("");
                that.controls.tableName.val(that.model.Name);
                that.controls.tableLocation.val(that.model.Position);
                that.controls.chartName.val("");
                that.controls.chartLocation.val("");
            }
            else if (that.model.SourceType == app.sourceTypes.chart) {
                that.controls.name.val("");
                that.controls.position.val("");
                that.controls.tableName.val("");
                that.controls.tableLocation.val("");
                that.controls.chartName.val(that.model.Name);
                var _p = that.utility.position(that.model.Position);
                that.controls.chartLocation.val(_p.sheet + " - " + _p.cell);
                that.controls.chartLocation.prop("address", that.model.Position);
            }
        }
        else {
            that.controls.name.val("");
            that.controls.position.val("");
            that.controls.tableName.val("");
            that.controls.tableLocation.val("");
            that.controls.chartName.val("");
            that.controls.chartLocation.val("");
            that.controls.chartLocation.prop("address", "");
        }

        if (!that.controls.main.hasClass("single") &&
            !that.controls.main.hasClass("chart") &&
            !that.controls.main.hasClass("table")) {
            that.controls.main.removeClass("manage add edit bulk single table chart clone step-first step-second step-third step-fourth")
                .addClass(that.model ? "add edit " +
                    (that.model.SourceType == app.sourceTypes.point ? "single" : (that.model.SourceType == app.sourceTypes.table ? "table" : "chart")) +
                    "" : (that.bulk ? "add bulk" : "add single"));
        }
        if (that.controls.main.hasClass("single")) {
            that.controls.sourceTypeNav.removeClass("is-selected");
            $(that.controls.sourceTypeNav[0]).addClass("is-selected");
            $(that.controls.back.find(".n-name-add")[0]).text("Add Source Point");
        }
        if (callback) {
            callback();
        }
    };

    that.utility = {
        model: function (id) {
            var _m = null;
            if (id && that.points) {
                var _t = [];
                $.each(that.points, function (i, d) {
                    _t.push(d.Id);
                });
                if ($.inArray(id, _t) > -1) {
                    _m = that.points[$.inArray(id, _t)];
                }
            }
            return _m;
        },
        mode: function (callback) {
            if (Office.context.document.mode == Office.DocumentMode.ReadOnly) {
                that.popup.message({ success: false, title: "Please click \"edit workbook\" button under the excel ribbon." });
            }
            else {
                callback();
            }
        },
        position: function (p) {
            if (p != null && p != undefined) {
                var _o = p.indexOf("'!"), _i = p.lastIndexOf(_o > -1 ? "'!" : "!"), _s = p.substr(0, _i).replace(new RegExp(/('')/g), '\''), _c = p.substr(_o > -1 ? _i + 2 : _i + 1, p.length);
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
        entered: function () {
            var name = $.trim(that.isPoint() || that.isBulk() ? that.controls.name.val() : (that.isTable() ? that.controls.tableName.val() : (that.isChart() ? that.controls.chartName.val() : ""))),
                position = $.trim(that.isPoint() || that.isBulk() ? that.controls.position.val() : (that.isTable() ? that.controls.tableLocation.val() : (that.isChart() ? that.controls.chartLocation.prop("address") : "")));
            return { name: name, position: position };
        },
        changed: function (callback) {
            var entered = that.utility.entered();
            if (that.model) {
                if (that.model.SourceType == app.sourceTypes.point) {
                    callback({ changed: !(that.model.NamePosition == entered.name && that.model.Position == entered.position) });
                }
                else {
                    callback({ changed: !(that.model.Name == entered.name && that.model.Position == entered.position) });
                }
            }
            else {
                callback({ changed: !(entered.name.length == 0 && entered.position.length == 0) });
            }
        },
        valid: function (options, callback) {
            var _p = that.utility.position(options.position);
            if ((!that.bulk && _p.cell.indexOf(":") == -1) || that.bulk || that.isTable() || that.isChart()) {
                if (that.isChart()) {
                    Excel.run(function (ctx) {
                        var _chart = ctx.workbook.worksheets.getItem(_p.sheet).charts.getItem(_p.cell);
                        var _image = _chart.getImage();
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded, data: _image.value });
                        });
                    }).catch(function (error) {
                        callback({ status: app.status.failed, message: "The selected " + options.title + " is invalid." });
                    });
                }
                else {
                    Excel.run(function (ctx) {
                        var r = ctx.workbook.worksheets.getItem(_p.sheet).getRange(_p.cell);
                        r.load("address,text");
                        return ctx.sync().then(function () {
                            if (!that.bulk) {
                                if (that.isPoint()) {
                                    callback({ status: app.status.succeeded, data: r.text[0][0] ? $.trim(r.text[0][0]) : "" });
                                }
                                else if (that.isTable()) {
                                    if (r.text.length > 1 && r.text[0].length > 1) {
                                        if (!that.utility.isEmpty({ data: r.text })) {
                                            var _image = r.getImage();
                                            return ctx.sync().then(function () {
                                                var _imageSrc = _image.value;
                                                that.range.formats({ range: r }, function (result) {
                                                    if (result.status == app.status.succeeded) {
                                                        callback({ status: app.status.succeeded, data: JSON.stringify({ image: _imageSrc, table: result.data }) });
                                                    }
                                                    else {
                                                        callback({ status: app.status.failed, message: "Get source table styles failed." });
                                                    }
                                                });
                                            });
                                        }
                                        else {
                                            callback({ status: app.status.failed, message: "The source table cannot be blank." });
                                        }
                                    }
                                    else {
                                        callback({ status: app.status.failed, message: "A table must contain at least 2 rows and 2 columns, please try again." });
                                    }
                                }
                            }
                            else {
                                if (!that.utility.isEmpty({ data: r.text })) {
                                    if (r.text.length > 0 && r.text[0].length == 2) {
                                        callback({ status: app.status.succeeded, data: r.text });
                                    }
                                    else {
                                        callback({ status: app.status.failed, message: "Only 2 adjacent columns can be selected." });
                                    }
                                }
                                else {
                                    callback({ status: app.status.failed, message: "The selected range cannot be blank." });
                                }
                            }
                        });
                    }).catch(function (error) {
                        callback({ status: app.status.failed, message: "The selected " + options.title + " position is invalid." });
                    });
                }
            }
            else {
                callback({ status: app.status.failed, message: "Only 1 cell can be selected for " + options.title + "." });
            }
        },
        exist: function (options, callback) {
            var a = [], existed = false;
            if (options.name) {
                $.each(that.points, function (i, d) {
                    a.push(d.Name);
                });
                if (that.model) {
                    existed = that.model.Name != options.name && $.inArray(options.name, a) > -1;
                }
                else {
                    existed = $.inArray(options.name, a) > -1;
                }
            }
            callback({ status: existed ? app.status.failed : app.status.succeeded, message: existed ? (options.name + " already exists, please select a different name.") : "" });
        },
        validation: function (callback) {
            var entered = that.utility.entered(),
                rangeId = that.model ? that.model.RangeId : app.guid(),
                nameRangeId = that.model ? that.model.NameRangeId : app.guid(),
                namePosition = entered.name,
                position = entered.position,
                success = true,
                values = [];
            if (!that.bulk && namePosition.length == 0) {
                success = false;
                values.push([that.isPoint() ? "Source Point Name" : (that.isTable() ? "Source Table Title" : "Source Chart Title")]);
            }
            if (position.length == 0) {
                success = false;
                values.push([(that.isPoint() || that.isBulk() || that.isTable()) ? "Select range" : "Select Chart"]);
            }
            if (!success) {
                callback({ status: app.status.failed, message: { success: success, title: "Please enter the following required fields:", values: values } });
            }
            else {
                if (!that.bulk) {
                    if (that.isPoint()) { // Source points
                        that.utility.valid({ position: namePosition, title: "source point name" }, function (result) {
                            if (result.status == app.status.succeeded) {
                                var name = result.data;
                                if (name.length > 0 && name.length <= 255) {
                                    that.utility.exist({ name: name }, function (result) {
                                        if (result.status == app.status.succeeded) {
                                            that.utility.valid({ position: position, title: "range" }, function (result) {
                                                if (result.status == app.status.succeeded) {
                                                    callback({
                                                        status: app.status.succeeded,
                                                        data: {
                                                            Id: that.model ? that.model.Id : "",
                                                            Name: name,
                                                            CatalogName: that.filePath,
                                                            DocumentId: that.documentId,
                                                            RangeId: rangeId,
                                                            NameRangeId: nameRangeId,
                                                            NamePosition: namePosition,
                                                            Position: position,
                                                            SourceType: 1,
                                                            Value: result.data
                                                        }
                                                    });
                                                }
                                                else {
                                                    callback({ status: app.status.failed, message: { success: false, title: result.message } });
                                                }
                                            });
                                        }
                                        else {
                                            callback({ status: app.status.failed, message: { success: false, title: result.message } });
                                        }
                                    });
                                }
                                else {
                                    callback({ status: app.status.failed, message: { success: false, title: name.length > 0 ? "The source point name cannot exceed 255 characters." : "The source point name cannot be blank." } });
                                }
                            }
                            else {
                                callback({ status: app.status.failed, message: { success: false, title: result.message } });
                            }
                        });
                    }
                    else { // Source table and chart
                        var name = namePosition;
                        if (name.length > 0 && name.length <= 255) {
                            that.utility.exist({ name: name }, function (result) {
                                if (result.status == app.status.succeeded) {
                                    that.utility.valid({ position: position, title: that.isTable() ? "table" : "chart" }, function (result) {
                                        if (result.status == app.status.succeeded) {
                                            callback({
                                                status: app.status.succeeded,
                                                data: {
                                                    Id: that.model ? that.model.Id : "",
                                                    Name: name,
                                                    CatalogName: that.filePath,
                                                    DocumentId: that.documentId,
                                                    RangeId: rangeId,
                                                    NameRangeId: null,
                                                    NamePosition: null,
                                                    Position: position,
                                                    SourceType: that.isTable() ? 2 : 3,
                                                    Value: result.data
                                                }
                                            });
                                        }
                                        else {
                                            callback({ status: app.status.failed, message: { success: false, title: result.message } });
                                        }
                                    });
                                }
                                else {
                                    callback({ status: app.status.failed, message: { success: false, title: result.message } });
                                }
                            });
                        }
                        else {
                            callback({ status: app.status.failed, message: { success: false, title: name.length > 0 ? "The source table title cannot exceed 255 characters." : "The source table title cannot be blank." } });
                        }
                    }
                }
                else {
                    that.utility.valid({ position: position, title: "range" }, function (result) {
                        if (result.status == app.status.succeeded) {
                            callback({
                                status: app.status.succeeded,
                                data: {
                                    Id: that.model ? that.model.Id : "",
                                    Name: "",
                                    CatalogName: that.filePath,
                                    DocumentId: that.documentId,
                                    RangeId: rangeId,
                                    NameRangeId: "",
                                    NamePosition: "",
                                    Position: position,
                                    SourceType: 1,
                                    Value: result.data
                                }
                            });
                        }
                        else {
                            callback({ status: app.status.failed, message: { success: false, title: result.message } });
                        }
                    });
                }
            }
        },
        index: function (options) {
            var _i = -1;
            $.each(that.points, function (m, n) {
                if (n.Id == options.Id) {
                    _i = m;
                    return false;
                }
            });
            return _i;
        },
        add: function (options) {
            options.PublishedStatus = true;
            options.NameExisted = true;
            options.ValueExisted = true;
            options.Existed = true;
            that.points.push(options);
        },
        remove: function (options) {
            that.points.splice(that.utility.index(options), 1);
        },
        update: function (options) {
            var _nx = that.points[that.utility.index(options)].NameExisted, _vx = that.points[that.utility.index(options)].ValueExisted;
            that.points[that.utility.index(options)] = options;
            that.points[that.utility.index(options)].PublishedStatus = options.Value == options.PublishedHistories[0].Value;
            that.points[that.utility.index(options)].NameExisted = _nx;
            that.points[that.utility.index(options)].ValueExisted = _vx;
            that.points[that.utility.index(options)].Existed = _nx && _vx;
        },
        updateAndAppendPublishHisotry: function (options) {
            var originalPublishHistories = that.points[that.utility.index(options)].PublishedHistories;
            that.utility.update(options);
            var latestPublishHistories = that.points[that.utility.index(options)].PublishedHistories;
            that.points[that.utility.index(options)].PublishedHistories = latestPublishHistories.concat(originalPublishHistories);
        },
        selected: function () {
            var _s = [];
            that.controls.list.find(".point-item .ckb-wrapper input").each(function (i, d) {
                if ($(d).prop("checked")) {
                    var _id = $(d).closest(".point-item").data("id"), _m = that.utility.model(_id);
                    _s.push({
                        SourcePointId: _id,
                        Name: app.string($.trim($(d).closest(".point-item").find(".i2 .s-name").text())),
                        RangeId: $(d).closest(".point-item").data("range"),
                        CurrentValue: _m.Value,
                        Position: $(d).closest(".point-item").data("position"),
                        NameRangeId: $(d).closest(".point-item").data("namerange"),
                        NamePosition: $(d).closest(".point-item").data("nameposition"),
                        SourceType: $(d).closest(".point-item").data("sourcetype")
                    });
                }
            });
            return _s;
        },
        all: function () {
            var _s = [];
            $.each(that.points, function (i, d) {
                _s.push({
                    SourcePointId: d.Id,
                    Name: app.string(d.Name),
                    RangeId: d.RangeId,
                    CurrentValue: app.string(d.Value),
                    Position: d.Position,
                    NameRangeId: d.NameRangeId,
                    NamePosition: d.NamePosition,
                    SourceType: d.SourceType
                });
            });
            return _s;
        },
        fileName: function (path) {
            return path.lastIndexOf("/") > -1 ? path.substr(path.lastIndexOf("/") + 1) : (path.lastIndexOf("\\") > -1 ? path.substr(path.lastIndexOf("\\") + 1) : path);
        },
        filePath: function (path, libraryPath) {
            return decodeURI(path).replace(decodeURI(libraryPath), "");
        },
        value: function (options, callback) {
            if (options.index < options.data.length) {
                that.range.exist({ Data: options.data[options.index], IsNameRange: true }, function (ret) {
                    that.range.exist({ Data: options.data[options.index], IsNameRange: false }, function (result) {
                        if (result.status == app.status.succeeded && ret.status == app.status.succeeded) {
                            options.data[options.index].DocumentValue = result.data.text;
                            options.data[options.index].DocumentPosition = result.data.address;
                            options.data[options.index].DocumentNameValue = ret.data.text;
                            options.data[options.index].DocumentNamePosition = ret.data.address;
                            options.data[options.index].Existed = true;
                        }
                        else {
                            options.data[options.index].Existed = false;
                        }
                        options.index++;
                        that.utility.value(options, callback);
                    });
                });
            }
            else {
                callback(options);
            }
        },
        dirty: function (options, callback) {
            var _f = false;
            $.each(options.data, function (i, d) {
                if (d.Existed && (d.CurrentValue != d.DocumentValue || d.Position != d.DocumentPosition || d.Name != d.DocumentNameValue || d.NamePosition != d.DocumentNamePosition)) {
                    _f = true;
                    return false;
                }
            });
            callback(_f);
        },
        pager: {
            init: function (options, callback) {
                that.controls.pagerValue.val("");
                that.controls.indexes.html("");
                that.pagerIndex = options.index && options.index > 0 ? options.index : 1;
                that.ui.list({ refresh: options.refresh }, callback);
            },
            prev: function () {
                that.controls.pagerValue.val("");
                that.utility.pager.updatePager();
                that.pagerIndex--;
                that.ui.list({ refresh: false });
            },
            next: function () {
                that.controls.pagerValue.val("");
                that.utility.pager.updatePager();
                that.pagerIndex++;
                that.ui.list({ refresh: false });
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
        addresses: function (options, callback) {
            if (options.index == undefined) {
                options.index = 0;
                options.position = that.utility.position(options.Position);
                options.cells = [];
                options.addresses = [];
                options.existed = [];
                options.texts = [];
                for (var i = 0; i < options.Value.length; i++) {
                    for (var m = 0; m < options.Value[i].length; m++) {
                        options.cells.push({ row: i, cell: m });
                    }
                }
            }
            if (options.index < options.cells.length / 2) {
                Excel.run(function (ctx) {
                    var r = ctx.workbook.worksheets.getItem(options.position.sheet).getRange(options.position.cell);
                    var c = r.getCell(options.cells[2 * options.index].row, options.cells[2 * options.index].cell);
                    c.load("text,address");
                    return ctx.sync().then(function () {
                        if (c.text && $.trim(c.text[0][0]) != "" && $.trim(c.text[0][0]).length <= 255) {
                            var _t = $.trim(c.text[0][0]), _a = c.address;
                            that.utility.exist({ name: _t }, function (result) {
                                if (result.status == app.status.succeeded) {
                                    var cc = r.getCell(options.cells[2 * options.index + 1].row, options.cells[2 * options.index + 1].cell);
                                    cc.load("text,address");
                                    return ctx.sync().then(function () {
                                        if ($.inArray(_t, options.texts) == -1) {
                                            options.addresses.push({ title: _t, text: $.trim(cc.text[0][0]), address: cc.address, nameAddress: _a });
                                            options.texts.push(_t);
                                        }
                                        else {
                                            options.existed.push(_t);
                                        }
                                        options.index++;
                                        that.utility.addresses(options, callback);
                                    });
                                }
                                else {
                                    options.existed.push(_t);
                                    options.index++;
                                    that.utility.addresses(options, callback);
                                }
                            });
                        }
                        else {
                            options.index++;
                            that.utility.addresses(options, callback);
                        }
                    });
                }).catch(function (error) {
                    options.index++;
                    that.utility.addresses(options, callback);
                });
            }
            else {
                callback(options);
            }
        },
        unSelectAll: function () {
            that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", false);
            that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
        },
        height: function () {

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
        },
        scrollTop: function () {
            that.controls.list.scrollTop(0);
        },
        selectedSourceType: function (sourceTypeClass) {
            var _i = $(sourceTypeClass).find("li.is-selected").index();
            if (_i == 0) {
                return app.sourceTypes.point;
            }
            else if (_i == 1) {
                return app.sourceTypes.chart;
            }
            else if (_i == 2) {
                return app.sourceTypes.table;
            }
            else {
                return app.sourceTypes.all;
            }
        },
        sortType: function () {
            if (that.controls.headerListPoints.find(".i2.sort-desc,.i2.sort-asc").length > 0) {
                return { sortType: app.sortTypes.name, sortOrder: that.controls.headerListPoints.find(".i2").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
            else if (that.controls.headerListPoints.find(".i3.sort-desc,.i3.sort-asc").length > 0) {
                return { sortType: app.sortTypes.status, sortOrder: that.controls.headerListPoints.find(".i3").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
            else if (that.controls.headerListPoints.find(".i4.sort-desc,.i4.sort-asc").length > 0) {
                return { sortType: app.sortTypes.value, sortOrder: that.controls.headerListPoints.find(".i4").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
            else {
                return { sortType: app.sortTypes.df };
            }
        },
        isEmpty: function (options) {
            var _flag = true;
            if (options.data && options.data.length > 0) {
                $.each(options.data, function (_a, _b) {
                    $.each(_b, function (_c, _d) {
                        if ($.trim(_d).length > 0) {
                            _flag = false;
                            return;
                        }
                    });
                });
            }
            return _flag;
        }
    };

    that.action = {
        body: function () {
            $(".search-tooltips").hide();
            that.controls.chartList.removeClass("active");
        },
        add: function () {
            that.utility.mode(function () {
                that.model = null;
                that.bulk = false;
                that.default();
                that.controls.name.focus();
            });
        },
        bulk: function () {
            that.utility.mode(function () {
                that.model = null;
                that.bulk = true;
                that.default();
                that.controls.name.focus();
            });
        },
        back: function () {
            if (that.controls.main.hasClass("clone")) {
                if (that.controls.main.hasClass("step-fourth")) {
                    //that.controls.cloneNext.addClass("disabled");
                    //that.controls.cloneBtn.addClass("disabled");
                    //that.controls.main.removeClass("step-fourth").addClass("step-third");
                    //that.action.clone.check();
                    //that.action.backToList();
                }
                else if (that.controls.main.hasClass("step-third")) {
                    that.controls.cloneNext.removeClass("disabled ms-Button--primary").addClass("ms-Button--primary");
                    that.controls.cloneBtn.removeClass("ms-Button--primary").addClass("disabled");
                    that.controls.main.removeClass("step-third").addClass("step-first");
                }
                else {
                    that.action.backToList();
                }
            }
            else {
                that.utility.changed(function (result) {
                    if (result.changed) {
                        that.popup.confirm({
                            title: "Do you want to save your changes?"
                        },
                        function () {
                            that.controls.popupMain.removeClass("message process confirm browse active");
                            that.controls.save.click();
                        }, function () {
                            that.controls.popupMain.removeClass("message process confirm browse active");
                            that.action.backToList();
                        });
                    }
                    else {
                        that.action.backToList();
                    }
                });
            }
        },
        backToList: function () {
            that.controls.main.removeClass("add edit bulk single table chart clone step-first step-second step-third step-fourth").addClass("manage");
            that.utility.scrollTop();
        },
        refresh: function () {
            that.utility.mode(function () {
                that.points = [];
                that.list({ refresh: true, index: that.pagerIndex }, function (result) {
                    if (result.status == app.status.failed) {
                        that.popup.message({ success: false, title: result.error.statusText });
                    }
                    else {
                        that.popup.message({ success: true, title: "Refresh all source points succeeded." }, function () { that.popup.hide(3000); });
                    }
                });
            });
        },
        select: function (options) {
            that.utility.mode(function () {
                that.range.select(function (result) {
                    if (result.status == app.status.succeeded) {
                        options.input.val(result.data);
                    }
                    else {
                        that.popup.message({ success: false, title: result.message });
                    }
                });
            });
        },
        save: function () {
            that.utility.mode(function () {
                that.popup.processing(true);
                var _isPoint = that.isPoint(), _isTable = that.isTable(), _isChart = that.isChart();
                if (that.model) {
                    that.utility.validation(function (result) {
                        if (result.status == app.status.succeeded) {
                            that.action.backToList();
                            if (_isPoint) {
                                that.range.del({ RangeId: result.data.NameRangeId }, function (ret) {
                                    if (ret.status == app.status.succeeded) {
                                        that.range.create({ Position: result.data.NamePosition, RangeId: result.data.NameRangeId }, function (ret) {
                                            if (ret.status == app.status.succeeded) {
                                                that.range.del({ RangeId: result.data.RangeId }, function (ret) {
                                                    if (ret.status == app.status.succeeded) {
                                                        that.range.create({ Position: result.data.Position, RangeId: result.data.RangeId }, function (ret) {
                                                            if (ret.status == app.status.succeeded) {
                                                                that.service.edit({ data: result.data }, function (result) {
                                                                    if (result.status == app.status.succeeded) {
                                                                        that.utility.update(result.data);
                                                                        that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                                                        that.popup.message({ success: true, title: "Source point update complete." }, function () { that.popup.hide(3000); });
                                                                    }
                                                                    else {
                                                                        that.popup.message({ success: false, title: "Edit source point failed." });
                                                                    }
                                                                });
                                                            }
                                                            else {
                                                                that.popup.message({ success: false, title: "Create range in Excel failed." });
                                                            }
                                                        });
                                                    }
                                                    else {
                                                        that.popup.message({ success: false, title: "Delete the previous range failed." });
                                                    }
                                                });
                                            }
                                            else {
                                                that.popup.message({ success: false, title: "Create source point name range in Excel failed." });
                                            }
                                        });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Delete the previous source point name range failed." });
                                    }
                                });
                            }
                            else if (_isTable) {
                                that.range.del({ RangeId: result.data.RangeId }, function (ret) {
                                    if (ret.status == app.status.succeeded) {
                                        that.range.create({ Position: result.data.Position, RangeId: result.data.RangeId }, function (ret) {
                                            if (ret.status == app.status.succeeded) {
                                                that.service.edit({ data: result.data }, function (result) {
                                                    if (result.status == app.status.succeeded) {
                                                        that.utility.update(result.data);
                                                        that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                                        that.popup.message({ success: true, title: "Source table update complete." }, function () { that.popup.hide(3000); });
                                                    }
                                                    else {
                                                        that.popup.message({ success: false, title: "Edit source table failed." });
                                                    }
                                                });
                                            }
                                            else {
                                                that.popup.message({ success: false, title: "Create table range in Excel failed." });
                                            }
                                        });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Delete the previous table range failed." });
                                    }
                                });
                            }
                            else if (_isChart) {
                                that.service.edit({ data: result.data }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        that.utility.update(result.data);
                                        that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                        that.popup.message({ success: true, title: "Source chart update complete." }, function () { that.popup.hide(3000); });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Edit source chart failed." });
                                    }
                                });
                            }
                        }
                        else {
                            that.popup.message(result.message);
                        }
                    });
                }
                else {
                    that.utility.validation(function (result) {
                        if (result.status == app.status.succeeded) {
                            that.action.backToList();
                            if (!that.bulk) {
                                if (_isPoint) {
                                    that.range.create({ Position: result.data.NamePosition, RangeId: result.data.NameRangeId }, function (ret) {
                                        if (ret.status == app.status.succeeded) {
                                            that.range.create({ Position: result.data.Position, RangeId: result.data.RangeId }, function (ret) {
                                                if (ret.status == app.status.succeeded) {
                                                    that.service.add({ data: result.data }, function (result) {
                                                        if (result.status == app.status.succeeded) {
                                                            that.utility.add(result.data);
                                                            that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                                            that.popup.message({ success: true, title: "Add new source point succeeded." }, function () { that.popup.hide(3000); });
                                                        }
                                                        else {
                                                            that.popup.message({ success: false, title: "Add source point failed." });
                                                        }
                                                    });
                                                }
                                                else {
                                                    that.popup.message({ success: false, title: "Create range in Excel failed." });
                                                }
                                            });
                                        }
                                        else {
                                            that.popup.message({ success: false, title: "Create source point name range in Excel failed." });
                                        }
                                    });
                                }
                                else if (_isTable) {
                                    that.range.create({ Position: result.data.Position, RangeId: result.data.RangeId }, function (ret) {
                                        if (ret.status == app.status.succeeded) {
                                            that.service.add({ data: result.data }, function (result) {
                                                if (result.status == app.status.succeeded) {
                                                    that.utility.add(result.data);
                                                    that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                                    that.popup.message({ success: true, title: "Add new source table succeeded." }, function () { that.popup.hide(3000); });
                                                }
                                                else {
                                                    that.popup.message({ success: false, title: "Add source table failed." });
                                                }
                                            });

                                        }
                                        else {
                                            that.popup.message({ success: false, title: "Create table range in Excel failed." });
                                        }
                                    });
                                }
                                else if (_isChart) {
                                    that.service.add({ data: result.data }, function (result) {
                                        if (result.status == app.status.succeeded) {
                                            that.utility.add(result.data);
                                            that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                            that.popup.message({ success: true, title: "Add new source chart succeeded." }, function () { that.popup.hide(3000); });
                                        }
                                        else {
                                            that.popup.message({ success: false, title: "Add source chart failed." });
                                        }
                                    });
                                }
                            }
                            else {
                                that.utility.addresses(result.data, function (rst) {
                                    rst.index = undefined;
                                    that.action.bulkAdd(rst, function (rt) {
                                        that.utility.pager.init({ refresh: false, index: that.pagerIndex });
                                        var ifBackToList = false;
                                        if (rt.status == app.status.succeeded) {
                                            if (rst.existed.length > 0) {
                                                var _v = [];
                                                $.each(rst.existed, function (_i, _d) {
                                                    _v.push([_d]);
                                                });
                                                that.popup.message({ success: false, title: "The following Source Points already exist, please input a unique name:", values: _v }, function () { that.popup.back(); });
                                            }
                                            else {
                                                that.popup.message({ success: true, title: rt.success + " Source Point(s) have been added to the catalog." }, function () {
                                                    that.popup.hide(3000);
                                                    that.action.back();
                                                    ifBackToList = true;
                                                });// Add bulk source points succeeded
                                            }
                                        }
                                        else {
                                            that.popup.message({ success: false, title: "Add " + rt.success + " source points succeeded, add " + rt.error + "source point failed." }, function () { that.popup.back(); });
                                        }
                                        if (ifBackToList === false) {
                                            that.utility.mode(function () {
                                                that.default();
                                            });
                                        }
                                    });
                                });
                            }
                        }
                        else {
                            that.popup.message(result.message);
                        }
                    });
                }
            });
        },
        bulkAdd: function (options, callback) {
            if (options.index == undefined) {
                options.index = 0;
                options.error = 0;
            }
            if (options.index < options.addresses.length) {
                var _i = options.addresses[options.index], _d = {
                    Id: "",
                    Name: _i.title,
                    CatalogName: options.CatalogName,
                    DocumentId: options.DocumentId,
                    RangeId: app.guid(),
                    NameRangeId: app.guid(),
                    NamePosition: _i.nameAddress,
                    Position: _i.address,
                    SourceType: 1,
                    Value: _i.text
                };

                that.range.create({ Position: _d.NamePosition, RangeId: _d.NameRangeId }, function (ret) {
                    if (ret.status == app.status.succeeded) {
                        that.range.create({ Position: _d.Position, RangeId: _d.RangeId }, function (ret) {
                            if (ret.status == app.status.succeeded) {
                                that.service.add({ data: _d }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        that.utility.add(result.data);
                                        options.index++;
                                        that.action.bulkAdd(options, callback);
                                    }
                                    else {
                                        options.index++;
                                        options.error++;
                                        that.action.bulkAdd(options, callback);
                                    }
                                });
                            }
                            else {
                                options.index++;
                                options.error++;
                                that.action.bulkAdd(options, callback);
                            }
                        });
                    }
                    else {
                        options.index++;
                        options.error++;
                        that.action.bulkAdd(options, callback);
                    }
                });
            }
            else {
                callback({ status: options.error > 0 ? app.status.failed : app.status.succeeded, error: options.error, success: options.addresses.length - options.error });
            }
        },
        del: function (i, o) {
            that.utility.mode(function () {
                var _sourceType = o.data("sourcetype"), _titleType = _sourceType == app.sourceTypes.point ? "point" : (_sourceType == app.sourceTypes.table ? "table" : "chart");
                that.popup.confirm({
                    title: "Do you want to delete the source " + _titleType + "?"
                }, function () {
                    that.popup.processing(true);
                    that.service.del({
                        Id: i
                    }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.popup.message({ success: true, title: "Delete source " + _titleType + " succeeded." }, function () { that.popup.hide(3000); });
                            that.utility.remove({ Id: i });
                            that.ui.remove({ Id: i });
                            that.utility.pager.init({ refresh: false, index: that.controls.list.find(".point-item").length > 0 ? that.pagerIndex : that.pagerIndex - 1 });
                        }
                        else {
                            that.popup.message({ success: false, title: "Delete source " + _titleType + " failed." });
                        }
                    });
                }, function () {
                    that.controls.popupMain.removeClass("message process confirm active");
                });
            });
        },
        deleteSelected: function () {
            var _s = that.utility.selected(), _ss = [], _st = [], _stp = "";
            if (_s && _s.length > 0) {
                $.each(_s, function (_y, _z) {
                    _ss.push(_z.SourcePointId);
                    if ($.inArray(_z.SourceType, _st) == -1) {
                        _st.push(_z.SourceType);
                    }
                });
                _st.sort();
                _stp = _st.join(",").replace(app.sourceTypes.point, "point").replace(app.sourceTypes.table, "table").replace(app.sourceTypes.chart, "chart");
                that.utility.mode(function () {
                    that.popup.confirm({
                        title: "Do you want to delete the selected source " + _stp + "?"
                    }, function () {
                        that.popup.processing(true);
                        that.service.deleteSelected({ data: { "": _ss } }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.popup.message({ success: true, title: "Delete source " + _stp + " succeeded." }, function () { that.popup.hide(3000); });
                                $.each(_ss, function (_m, _n) {
                                    that.utility.remove({ Id: _n });
                                    that.ui.remove({ Id: _n });
                                });
                                that.utility.unSelectAll();
                                that.utility.pager.init({ refresh: false, index: that.controls.list.find(".point-item").length > 0 ? that.pagerIndex : that.pagerIndex - 1 });
                            }
                            else {
                                that.popup.message({ success: false, title: "Delete source " + _stp + " failed." });
                            }
                        });
                    }, function () {
                        that.controls.popupMain.removeClass("message process confirm active");
                    });
                });
            }
            else {
                that.popup.message({ success: false, title: "Please select source point, table or chart." });
            }
        },
        edit: function (i, o) {
            that.utility.mode(function () {
                that.bulk = false;
                that.model = that.utility.model(i);
                if (that.model) {
                    if (that.model.SourceType == app.sourceTypes.point) {
                        $(that.controls.back.find(".n-name-edit")[0]).text("Edit Source Point");
                        that.model.NameRangeId = that.model.NameRangeId && that.model.NameRangeId != null ? that.model.NameRangeId : "";
                        that.model.NamePosition = that.model.NamePosition && that.model.NamePosition != null ? that.model.NamePosition : "";
                        that.range.goto({ RangeId: that.model.NameRangeId }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.model.NamePosition = result.data.address;
                                that.range.goto({ RangeId: that.model.RangeId }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        that.model.Position = result.data.address;
                                        that.default(function () { that.action.associated(); });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "The range in the Excel has been deleted." });
                                    }
                                });
                            }
                            else {
                                that.popup.message({ success: false, title: "The source point name range in the Excel has been deleted." });
                            }
                        });
                    }
                    else if (that.model.SourceType == app.sourceTypes.table) {
                        $(that.controls.back.find(".n-name-edit")[0]).text("Edit Source Table");
                        that.range.goto({ RangeId: that.model.RangeId }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.model.Position = result.data.address;
                                that.default(function () { that.action.associated(); });
                            }
                            else {
                                that.popup.message({ success: false, title: "The table range in the Excel has been deleted." });
                            }
                        });
                    }
                    else if (that.model.SourceType == app.sourceTypes.chart) {
                        $(that.controls.back.find(".n-name-edit")[0]).text("Edit Source Chart");
                        that.default(function () { that.action.associated(); });
                    }
                }
                else {
                    that.popup.message({ success: false, title: "The source " + app.sourceTitle(o.data("sourcetype")) + " has been deleted." });
                }
            });
        },
        history: function (o) {
            o.hasClass("item-more") ? o.removeClass("item-more") : o.addClass("item-more");
        },
        // The published item is equal to the document item(compare name & value's text and position)
        isEqualPublish: function (o) {
            // Get the published item from points
            var _i = that.utility.index({ Id: o.SourcePointId });
            if (_i < 0) {
                return false;
            }
            var _pointItem = that.points[_i];
            var _publishedItem = _pointItem;
            if (_pointItem.PublishedHistories && _pointItem.PublishedHistories.length > 0) {
                _publishedItem = _pointItem.PublishedHistories[0];
            }
            if (_publishedItem.Name == o.DocumentNameValue &&
                _pointItem.NamePosition == o.DocumentNamePosition &&
                _publishedItem.Value == o.DocumentValue &&
                _publishedItem.Position == o.DocumentPosition) {
                return true;
            }
            else {
                return false;
            }
        },
        publish: function () {
            var _s = that.utility.selected();
            if (_s && _s.length > 0) {
                that.popup.processing(true);
                that.utility.value({ index: 0, data: _s }, function (result) {
                    var _ss = [], _sf = false, _countNoChange = 0;
                    $.each(result.data, function (_m, _n) {
                        if (_n.Existed) {
                            if (that.action.isEqualPublish(_n)) {
                                // No change on current item
                                _countNoChange++;
                            }
                            else {
                                // Changed on current item
                                _ss.push(_n);
                            }
                        }
                        else {
                            _sf = true;
                        }
                    });
                    if (_ss.length > 0) {
                        that.utility.dirty({ data: _ss }, function (f) {
                            if (f) {
                                that.popup.confirm({
                                    title: "The cell value or location is not the same with the one in Add-in. Please click 'No' to cancel it and click 'refresh' to manually sync the value or location from excel to Add-in. Are you sure to publish the selected source point still?"
                                },
                                    function () {
                                        that.popup.message({ success: false, title: "Publish Cancelled: No changes detected, click the refresh button and try again." });
                                        /*that.popup.processing(true);
                                        that.service.publish({ data: { "": _ss } }, function (result) {
                                            if (result.status == app.status.succeeded) {
                                                $.each(result.data.SourcePoints, function (_i, _d) {
                                                    that.utility.update(_d);
                                                });
                                                that.ui.publish(result.data, function () {
                                                    that.popup.message({ success: _sf ? false : true, title: _sf ? "Not all Source Points were published, please check Source Point list for errors." : "Publish source point succeeded." }, function () { that.popup.hide(3000); });
                                                });
                                            }
                                            else {
                                                that.popup.message({ success: false, title: "Publish source point failed." });
                                            }
                                        });*/
                                    },
                                function () {
                                    that.controls.popupMain.removeClass("message process confirm active");
                                });
                            }
                            else {
                                that.service.publish({ data: { "": _ss } }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        $.each(result.data.SourcePoints, function (_i, _d) {
                                            that.utility.updateAndAppendPublishHisotry(_d);
                                        });
                                        that.ui.publish(result.data, function () {
                                            that.popup.message({ success: _sf ? false : true, title: _sf ? "Not all Source Points were published, please check Source Point list for errors." : "Publish source point succeeded." }, function () {
                                                that.popup.hide(3000);
                                            });
                                        });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Publish source point failed." });
                                    }
                                });
                            }
                        });
                    }
                    else {
                        if (_countNoChange > 0) {
                            that.popup.message({ success: false, title: "Publish Cancelled: No changes detected, click the refresh button and try again." });
                        }
                        else {
                            that.popup.message({ success: false, title: "The source points you selected have been deleted." });
                        }
                    }
                });
            }
            else {
                that.popup.message({ success: false, title: "Please select source point, table or chart." });
            }
        },
        publishAll: function () {
            var _s = that.utility.all();
            if (_s && _s.length > 0) {
                that.popup.processing(true);
                that.utility.value({ index: 0, data: _s }, function (result) {
                    var _ss = [], _sf = false, _countNoChange = 0;
                    $.each(result.data, function (_m, _n) {
                        if (_n.Existed) {
                            if (that.action.isEqualPublish(_n)) {
                                // No change on current item
                                _countNoChange++;
                            }
                            else {
                                // Changed on current item
                                _ss.push(_n);
                            }
                        }
                        else {
                            _sf = true;
                        }
                    });
                    if (_ss.length > 0) {
                        that.utility.dirty({ data: _ss }, function (f) {
                            if (f) {
                                that.popup.confirm({
                                    title: "The cell value or location is not the same with the one in Add-in. Please click 'No' to cancel it and click 'refresh' to manually sync the value or location from excel to Add-in. Are you sure to publish the selected source point still?"
                                },
                                    function () {
                                        that.popup.message({ success: false, title: "Publish Cancelled: No changes detected, click the refresh button and try again." });
                                        /*that.popup.processing(true);
                                        that.service.publish({ data: { "": _ss } }, function (result) {
                                            if (result.status == app.status.succeeded) {
                                                $.each(result.data.SourcePoints, function (_i, _d) {
                                                    that.utility.update(_d);
                                                });
                                                that.ui.publish(result.data, function () {
                                                    that.popup.message({ success: _sf ? false : true, title: _sf ? "Not all Source Points were published, please check Source Point list for errors." : "Publish source point succeeded." }, function () {
                                                        that.popup.hide(3000);
                                                    });
                                                });
                                            }
                                            else {
                                                that.popup.message({ success: false, title: "Publish source point failed." });
                                            }
                                        });*/
                                    },
                                function () {
                                    that.controls.popupMain.removeClass("message process confirm active");
                                });
                            }
                            else {
                                that.service.publish({ data: { "": _ss } }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        $.each(result.data.SourcePoints, function (_i, _d) {
                                            that.utility.updateAndAppendPublishHisotry(_d);
                                        });
                                        that.ui.publish(result.data, function () {
                                            that.popup.message({ success: _sf ? false : true, title: _sf ? "Not all Source Points were published, please check Source Point list for errors." : "Publish source point succeeded." }, function () {
                                                that.popup.hide(3000);
                                            });
                                        });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Publish source point failed." });
                                    }
                                });
                            }
                        });
                    }
                    else {
                        if (_countNoChange > 0) {
                            that.popup.message({ success: false, title: "Publish Cancelled: No changes detected, click the refresh button and try again." });
                        }
                        else {
                            that.popup.message({ success: false, title: "All source points have been deleted." });
                        }
                    }
                });
            }
            else {
                that.popup.message({ success: false, title: "There is no source point." });
            }
        },
        goto: function (i, o) {
            var _selectedItem = o.hasClass("point-item") ? o : o.closest(".point-item");
            that.controls.list.find(".point-item.selected").removeClass("selected");
            _selectedItem.addClass("selected");
            that.utility.mode(function () {
                var _m = that.utility.model(i);
                if (_m) {
                    if (_m.SourceType == app.sourceTypes.point || _m.SourceType == app.sourceTypes.table) {
                        that.range.exist({ Data: _m, IsNameRange: true }, function (result) {
                            if (result.status == app.status.failed) {
                                that.popup.message({ success: false, title: "The source point name range in the Excel has been deleted." });
                            }
                            else {
                                that.range.goto({ RangeId: _m.RangeId }, function (result) {
                                    if (result.status == app.status.failed) {
                                        that.popup.message({ success: false, title: "The range in the Excel has been deleted." });
                                    }
                                });
                            }
                        });
                    }
                    else {
                        that.range.exist({ Data: _m, IsNameRange: false }, function (result) {
                            if (result.status == app.status.failed) {
                                that.popup.message({ success: false, title: "The source chart in the Excel has been deleted." });
                            }
                            else {
                                that.range.gotoWorkSheet({ Sheet: that.utility.position(result.data.address).sheet }, function (result) {
                                    if (result.status == app.status.failed) {
                                        that.popup.message({ success: false, title: "The worksheet in the Excel has been deleted." });
                                    }
                                });
                            }
                        });
                    }
                }
                else {
                    that.popup.message({ success: false, title: "The source " + app.sourceTitle(o.data("sourcetype")) + " has been deleted." });
                }
            });
        },
        up: function () {
            if (!that.controls.moveUp.hasClass("disabled")) {
                that.controls.moveDown.removeClass("disabled");
                var _i = that.controls.list.find(".point-item.selected").index();
                if (_i <= 1) {
                    that.controls.moveUp.removeClass("disabled").addClass("disabled");
                }
                _i--;

                if (_i < 0) {
                    _i = 0;
                }
                if (_i >= 0) {
                    that.action.goto((that.controls.list.find(">li").eq(_i)).data("id"), that.controls.list.find(">li").eq(_i));
                }
            }
        },
        down: function () {
            if (!that.controls.moveDown.hasClass("disabled")) {
                that.controls.moveUp.removeClass("disabled");
                var _i = that.controls.list.find(".point-item.selected").index(), _l = that.controls.list.find(">li").length;
                if (_i >= _l - 2) {
                    that.controls.moveDown.removeClass("disabled").addClass("disabled");
                }
                _i++;

                if (_i >= _l) {
                    _i = _l - 1;
                }
                if (_i < _l) {
                    that.action.goto((that.controls.list.find(">li").eq(_i)).data("id"), that.controls.list.find(">li").eq(_i));
                }
            }
        },
        ok: function () {
            that.controls.popupMain.removeClass("active message process confirm");
        },
        associated: function () {
            that.controls.listAssociated.html("");
            that.service.associated(that.model, function (result) {
                if (result.status == app.status.succeeded) {
                    that.ui.associated({ data: result.data });
                }
                else {
                    that.popup.message({ success: false, title: "Get associated files failed." });
                }
            });
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
            that.utility.pager.init({ refresh: false });
        },
        clone: {
            source: null,
            destination: null,
            checkResult: null,
            init: function () {
                that.action.clone.source = null;
                that.action.clone.destination = null;
                that.action.clone.checkResult = null;
                that.controls.sourceFolder.val("");
                that.controls.destinationFolder.val("");
                var _s = $.trim(that.controls.sourceFolder.val()), _sd = that.controls.sourceFolder.data("default"),
                    _d = $.trim(that.controls.destinationFolder.val()), _dd = that.controls.destinationFolder.data("default");
                if (_s == "" || _s == _sd) {
                    that.controls.sourceFolder.val(_sd).addClass("input-default");
                }
                if (_d == "" || _d == _dd) {
                    that.controls.destinationFolder.val(_dd).addClass("input-default");
                }
                that.controls.cloneNext.removeClass("ms-Button--primary").addClass("disabled");
                that.controls.cloneResult.removeClass("clone-success clone-error");
                that.controls.main.removeClass("manage add edit bulk single table chart clone step-first step-second step-third step-fourth").addClass("add clone step-first");
            },
            next: function () {
                if (!that.controls.cloneNext.hasClass("disabled")) {
                    if (that.controls.main.hasClass("step-first")) {
                        that.controls.main.removeClass("step-first").addClass("step-third");
                        that.controls.cloneBtn.removeClass("ms-Button--primary").addClass("disabled");
                        that.action.clone.check();
                    }
                }
            },
            select: function (options) {
                if (that.controls.main.find(".clone-point-first").hasClass("source")) {
                    that.controls.sourceFolder.val(options.name).removeClass("input-default");
                    that.action.clone.source = options;
                    that.controls.sourceFolderName.text(options.name);
                }
                else if (that.controls.main.find(".clone-point-first").hasClass("dest")) {
                    that.controls.destinationFolder.val(options.name).removeClass("input-default");
                    that.action.clone.destination = options;
                    that.controls.destFolderName.text(options.name);
                }
                if (that.action.clone.source && that.action.clone.destination) {
                    that.controls.cloneNext.removeClass("disabled ms-Button--primary").addClass("ms-Button--primary");
                }
            },
            getDocumentId: function (options) {
                var _di = "";
                $.each(options.data, function (_i, _d) {
                    if (decodeURI(_d.Url).toUpperCase() == decodeURI(options.Url).toUpperCase()) {
                        _di = _d.DocumentId;
                    }
                });
                return _di;
            },
            check: function () {
                that.controls.listWillClone.html("");
                that.controls.listWillNotClone.html("");
                that.popup.processing(true);
                that.service.filesInFolder(that.action.clone.source, function (result) {
                    if (result.status == app.status.succeeded) {
                        var _j = [];
                        $.each(result.data.value, function (_i, _d) {
                            if (_d.FileSystemObjectType == "0") {
                                _j.push({
                                    Name: _d.FileLeafRef,
                                    Url: _d.EncodedAbsUrl,
                                    DocumentId: _d.OData__dlc_DocId ? _d.OData__dlc_DocId : "",
                                    IsExcel: _d.FileLeafRef.toUpperCase().indexOf(".XLSX") > -1,
                                    IsWord: _d.FileLeafRef.toUpperCase().indexOf(".DOCX") > -1,
                                    DestinationFileUrl: encodeURI(decodeURI(that.action.clone.destination.url) + "/" + decodeURI(_d.FileLeafRef)),
                                    DestinationFileDocumentId: "",
                                    Clone: false
                                });
                            }
                        });
                        if (_j.length > 0) {
                            that.service.checkCloneStatus({ data: { "": _j } }, function (result) {
                                if (result.status == app.status.succeeded) {
                                    that.action.clone.checkResult = result.data;
                                    that.ui.clone({ data: result.data });
                                    that.popup.processing(false);
                                }
                                else {
                                    that.popup.message({ success: false, title: "Check the clone file status failed." });
                                }
                            });
                        }
                        else {
                            that.popup.message({ success: false, title: "No files found in source folder." });
                        }
                    }
                    else {
                        that.popup.message({ success: false, title: "Get files in source folder failed." });
                    }
                });
            },
            copy: function (options, callback) {
                if (options.index == undefined) {
                    var _td = [];
                    $.each(options.data, function (_i, _d) {
                        if (_d.Clone) {
                            _td.push(_d);
                        }
                    });
                    options.data = _td;
                    options.index = 0;
                    that.popup.processing(true);
                }

                if (options.index < options.data.length) {
                    var _ci = options.data[options.index];
                    that.service.getFile(_ci, function (result) {
                        if (result.status == app.status.failed) {
                            _ci.Clone = false;
                            options.index++;
                            that.action.clone.copy(options, callback);
                        }
                        else {
                            that.service.copyFile({ Name: _ci.Name, Id: result.data.id }, function (result) {
                                _ci.Clone = result.status == app.status.succeeded;
                                options.index++;
                                that.action.clone.copy(options, callback);
                            });
                        }
                    });
                }
                else {
                    that.service.filesInFolder(that.action.clone.destination, function (result) {
                        if (result.status == app.status.succeeded) {
                            var _j = [];
                            $.each(result.data.value, function (_i, _d) {
                                if (_d.FileSystemObjectType == "0") {
                                    _j.push({
                                        Url: _d.EncodedAbsUrl,
                                        DocumentId: _d.OData__dlc_DocId ? _d.OData__dlc_DocId : "",
                                    });
                                }
                            });

                            $.each(options.data, function (_i, _d) {
                                if (_d.Clone) {
                                    _d.DestinationFileDocumentId = that.action.clone.getDocumentId({ Url: _d.DestinationFileUrl, data: _j });
                                }
                            });

                            callback({ data: options.data });
                        }
                        else {
                            that.popup.message({ success: false, title: "Get files in destination folder failed." });
                        }
                    });
                }
            },
            save: function (options, callback) {
                that.action.clone.copy({ data: $.extend([], that.action.clone.checkResult) }, function (result) {
                    that.service.clone({ data: { "": result.data } }, function (ret) {
                        if (ret.status == app.status.succeeded) {
                            that.ui.cloneResult(result);
                            that.popup.processing(false);
                        }
                        else {
                            that.popup.message({ success: false, title: "Clone folder failed." });
                        }
                    });
                });
            }
        },
        newFolder: {
            toggle: function (_show) {
                if (_show && that.controls.main.hasClass("step-first")) {
                    that.controls.popupBrowse.addClass("dplibrary");
                    if (that.controls.popupBrowse.hasClass("dpdeslibrary")) {
                        that.controls.popupBrowse.addClass("nf");
                    }
                    else {
                        that.controls.popupBrowse.removeClass("nf");
                    }
                }
                else {
                    that.controls.popupBrowse.removeClass("dplibrary");
                }
            },
            init: function () {
                that.controls.popupBrowseBack.hide()
                that.controls.popupBrowse.addClass("newfolder");
                that.controls.popupBrowseOKWrap.hide();
                that.controls.newFolderName.val("");
                that.action.newFolder.message({ message: "" });
            },
            create: function () {
                that.action.newFolder.message({ message: "" });
                var _name = $.trim(that.controls.newFolderName.val());
                if (_name.length > 0) {
                    that.browse.popup.processing(true);
                    var _p = that.browse.path[that.browse.path.length - 1];
                    if (_p.type == "library") {
                        that.service.newFolder({ siteId: _p.site, listId: _p.id, itemId: null, name: _name }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.action.newFolder.appendNewFolder({ type: "folder", name: decodeURI(result.data.name), id: result.data.id, siteId: _p.site, siteUrl: _p.siteUrl, listId: _p.id, url: result.data.webUrl, listName: _p.listName }, function () {
                                    that.action.newFolder.cancel();
                                });
                            }
                            else {
                                that.browse.popup.processing(false);
                                that.action.newFolder.message({ message: "Create folder failed: A file name can't contain any of the following characters:<br/>\/:*?\"<>|" });
                            }
                        });
                    }
                    else if (_p.type == "folder") {
                        that.service.newFolder({ siteId: _p.site, listId: _p.list, itemId: _p.id, name: _name }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.action.newFolder.appendNewFolder({ type: "folder", name: decodeURI(result.data.name), id: result.data.id, siteId: _p.site, siteUrl: _p.siteUrl, listId: _p.list, url: result.data.webUrl, listName: _p.listName }, function () {
                                    that.action.newFolder.cancel();
                                });
                            }
                            else {
                                that.browse.popup.processing(false);
                                that.action.newFolder.message({ message: "Create folder failed: A file name can't contain any of the following characters:<br/>\/:*?\"<>|" });
                            }
                        });
                    }
                }
                else {
                    that.action.newFolder.message({ message: "Please enter folder name." });
                }
            },
            message: function (options, callback) {
                that.controls.newFolderMessage.html(options.message);
            },
            cancel: function () {
                that.controls.popupBrowse.removeClass("newfolder");
                that.browse.popup.processing(false);
                that.browse.popup.nav();
                that.browse.buttonStatus();
            },
            appendNewFolder: function (options, callback) {
                var _d = [], _df = [], _dt = [];
                that.controls.popupBrowseList.find("li").each(function (_i, _e) {
                    if ($(_e).data("type") == "folder") {
                        _df.push({
                            type: $(_e).data("type"),
                            name: $(_e).text(),
                            id: $(_e).data("id"),
                            siteId: $(_e).data("site"),
                            listId: $(_e).data("list"),
                            url: $(_e).data("url"),
                            siteUrl: $(_e).data("siteurl"),
                            listName: $(_e).data("listname")
                        });
                    }
                    else {
                        _dt.push({
                            type: $(_e).data("type"),
                            name: $(_e).text(),
                            id: $(_e).data("id"),
                            siteId: $(_e).data("site"),
                            listId: $(_e).data("list"),
                            url: $(_e).data("url"),
                            siteUrl: $(_e).data("siteurl"),
                            listName: $(_e).data("listname")
                        });
                    }
                });
                _df.push(options);
                _df.sort(function (_a, _b) {
                    return (_a.name.toUpperCase() > _b.name.toUpperCase()) ? 1 : (_a.name.toUpperCase() < _b.name.toUpperCase()) ? -1 : 0;
                });
                _dt.sort(function (_a, _b) {
                    return (_a.name.toUpperCase() > _b.name.toUpperCase()) ? 1 : (_a.name.toUpperCase() < _b.name.toUpperCase()) ? -1 : 0;
                });
                _d = _df.concat(_dt);
                that.controls.popupBrowseList.html("");
                $.each(_d, function (i, d) {
                    var _h = "";
                    if (d.type == "folder") {
                        _h = '<li class="i-folder" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="folder" data-siteurl="' + d.siteUrl + '" data-listname="' + d.listName + '">' + (d.id == options.id ? "<i class=\"checked\"></i>" : "<i></i>") + d.name + '</li>';
                    }
                    else {
                        _h = '<li class="i-file i-clonefile" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="file" data-siteurl="' + d.siteUrl + '" data-listname="' + d.listName + '">' + d.name + '</li>';
                    }
                    that.controls.popupBrowseList.append(_h);
                });
                callback();
            }
        },
        sort: function (elem) {
            if (elem) {
                var _sd = elem.hasClass("sort-desc");
                that.controls.headerListPoints.find(".i2.sort-desc,.i2.sort-asc,.i3.sort-desc,.i3.sort-asc,.i4.sort-desc,.i4.sort-asc").removeClass("sort-asc sort-desc");
                elem.addClass(_sd ? "sort-asc" : "sort-desc");
            }
            else {
                that.controls.headerListPoints.find(".i2.sort-desc,.i2.sort-asc,.i3.sort-desc,.i3.sort-asc,.i4.sort-desc,.i4.sort-asc").removeClass("sort-asc sort-desc");
            }
            that.ui.list({ refresh: false });
        },
        resort: function () {
            that.controls.headerListPoints.find(".i2.sort-desc,.i2.sort-asc,.i3.sort-desc,.i3.sort-asc,.i4.sort-desc,.i4.sort-asc").removeClass("sort-asc sort-desc");
            that.controls.headerListPoints.find(".i2").addClass("sort-asc");
            that.ui.list({ refresh: false });
        },
        charts: function () {
            if (!that.controls.chartList.hasClass("active")) {
                that.range.workSheets(function (result) {
                    if (result.status == app.status.succeeded) {
                        that.range.charts({ sheets: result.data }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.ui.charts({ charts: result.data }, function () {
                                    that.controls.chartList.addClass("active");
                                });
                            }
                            else {
                                that.popup.message({ success: false, title: result.message });
                            }
                        });
                    }
                    else {
                        that.popup.message({ success: false, title: result.message });
                    }
                });
            }
            else {
                that.controls.chartList.removeClass("active");
            }
        },
        selectChart: function (o) {
            if ($.trim(that.controls.chartName.val()) == "") {
                that.controls.chartName.val(o.data("chart"));
            }
            that.controls.chartLocation.prop("address", ("'" + o.data("sheet") + "'" + "!" + o.data("chart"))).val($.trim(o.text()));
            that.controls.chartList.removeClass("active");
        }
    };

    that.browse = {
        path: [],
        currentSite: { values: [], webUrls: [] },
        browseFile: false,
        init: function () {
            that.api.token = "";
            that.browse.currentSite = { values: [], webUrls: [] };
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
            that.action.newFolder.toggle(false);
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
                        else {
                            that.browse.currentSite.values.push(result.data.id);
                            that.browse.currentSite.webUrls.push(result.data.webUrl);
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
            that.action.newFolder.toggle(false);
            that.service.sites(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _s = [];
                    $.each(result.data.value, function (i, d) {
                        _s.push({ id: d.id, name: d.name, type: "site", siteUrl: d.webUrl });
                    });
                    that.browse.libraries({ siteId: options.siteId, siteUrl: options.siteUrl, sites: _s });
                }
                else {
                    if (that.browse.currentSite.values.length > 0) {
                        that.browse.sites({ siteId: that.browse.currentSite.values[0], siteUrl: that.browse.currentSite.webUrls[0] });
                    }
                    else {
                        that.browse.popup.message("Get sites failed.");
                    }
                }
            });
        },
        libraries: function (options) {
            that.action.newFolder.toggle(false);
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
                                if (!(!that.browse.browseFile && that.action.clone.source != null && that.action.clone.source.url.toUpperCase() == decodeURI(_u).toUpperCase())) {
                                    _fd.push({ id: d.id, name: _nu, type: "folder", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                }
                            }
                            else if (d.file) {
                                if (that.browse.browseFile) {
                                    if (_n.toUpperCase().indexOf(".XLSX") > 0) {
                                        _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                    }
                                }
                                else {
                                    _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                }
                            }
                        });
                        _fi.sort(function (_a, _b) {
                            return (_a.name.toUpperCase() > _b.name.toUpperCase()) ? 1 : (_a.name.toUpperCase() < _b.name.toUpperCase()) ? -1 : 0;
                        });
                        that.browse.display({ data: _fd.concat(_fi) });
                        that.action.newFolder.toggle(true);
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
                                if (!(!that.browse.browseFile && that.action.clone.source != null && that.action.clone.source.url.toUpperCase() == decodeURI(_u).toUpperCase())) {
                                    _fd.push({ id: d.id, name: _nu, type: "folder", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                }
                            }
                            else if (d.file) {
                                if (that.browse.browseFile) {
                                    if (_n.toUpperCase().indexOf(".XLSX") > 0) {
                                        _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                    }
                                }
                                else {
                                    _fi.push({ id: d.id, name: _nu, type: "file", url: _u, siteId: options.siteId, siteUrl: options.siteUrl, listId: options.listId, listName: options.listName });
                                }
                            }
                        });
                        _fi.sort(function (_a, _b) {
                            return (_a.name.toUpperCase() > _b.name.toUpperCase()) ? 1 : (_a.name.toUpperCase() < _b.name.toUpperCase()) ? -1 : 0;
                        });
                        that.browse.display({ data: _fd.concat(_fi) });
                        that.action.newFolder.toggle(true);
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
                        that.files($.extend({}, { documentId: _d }, options));
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
                    _h = '<li class="i-folder" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="folder" data-siteurl="' + d.siteUrl + '" data-listname="' + d.listName + '">' + (that.browse.browseFile ? "" : "<i></i>") + d.name + '</li>';
                }
                else if (d.type == "file") {
                    _h = '<li class="i-file' + (that.browse.browseFile ? "" : " i-clonefile") + '" data-id="' + d.id + '" data-site="' + d.siteId + '" data-list="' + d.listId + '" data-url="' + d.url + '" data-type="file" data-siteurl="' + d.siteUrl + '" data-listname="' + d.listName + '">' + d.name + '</li>';
                }
                that.controls.popupBrowseList.append(_h);
            });
            if (options.data.length == 0) {
                that.controls.popupBrowseList.html("No items found.");
            }
            that.browse.popup.processing(false);
        },
        select: function (elem) {
            that.controls.popupBrowseOKWrap.hide();
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
                if (that.browse.browseFile) {
                    that.browse.popup.processing(true);
                    that.browse.file({ siteUrl: $(elem).data("siteurl"), listName: $(elem).data("listname"), name: $.trim($(elem).text()), url: that.browse.path[that.browse.path.length - 1].url + "/" + encodeURI($(elem).text()), fileName: $.trim($(elem).text()) });
                }
            }
        },
        checked: function (elem) {
            if ($(elem).find("i").hasClass("checked")) {
                $(elem).find("i").removeClass("checked");
                $(elem).parent().find("li.i-folder i").removeClass("checked");
            }
            else {
                $(elem).parent().find("li.i-folder i").removeClass("checked");
                $(elem).find("i").addClass("checked");
            }
            that.browse.buttonStatus();
        },
        buttonStatus: function () {
            if (that.controls.popupBrowseList.find("li.i-folder i.checked").length == 1) {
                that.controls.popupBrowseOKWrap.show();
            }
            else {
                that.controls.popupBrowseOKWrap.hide();
            }
        },
        selectFolder: function () {
            var elem = $(that.controls.popupBrowseList.find("li.i-folder i.checked")[0]).parent();
            that.browse.popup.hide();
            that.action.clone.select({
                id: $.trim($(elem).data("id")),
                name: $.trim($(elem).text()),
                url: $.trim(decodeURI($(elem).data("url"))),
                siteUrl: $.trim(decodeURI($(elem).data("siteurl"))),
                siteId: $.trim($(elem).data("site")),
                listName: $.trim(decodeURI($(elem).data("listname"))),
                listId: $.trim($(elem).data("list"))
            });
        },
        popup: {
            dft: function () {
                that.controls.popupBrowseOKWrap.hide();
                that.controls.popupBrowseList.html("");
                that.controls.popupBrowseBack.hide();
                that.controls.popupBrowseMessage.html("").hide();
                that.controls.popupBrowseLoading.hide();
                if (that.browse.browseFile) {
                    that.controls.popupBrowse.removeClass("clone-folder");
                }
                else {
                    that.controls.popupBrowse.addClass("clone-folder");
                }
            },
            show: function () {
                that.controls.popupMain.removeClass("message process confirm").addClass("active browse");
            },
            hide: function () {
                that.controls.popupBrowse.removeClass("dplibrary dpdeslibrary");
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
                that.controls.popupBrowseOKWrap.hide();
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

    that.document = {
        init: function (callback) {
            that.popup.processing(true);
            that.service.token({ endpoint: that.endpoints.token }, function (result) {
                if (result.status == app.status.succeeded) {
                    that.api.token = result.data;
                    that.api.host = that.utility.path()[0].toLowerCase();
                    that.service.token({ endpoint: that.endpoints.sharePointToken }, function (result) {
                        if (result.status == app.status.succeeded) {
                            that.api.sharePointToken = result.data;
                            that.document.site(null, callback);
                        }
                        else {
                            that.document.error({ title: "Get sharepoint access token failed." });
                        }
                    });
                }
                else {
                    that.document.error({ title: "Get graph access token failed." });
                }
            });
        },
        site: function (options, callback) {
            if (options == null) {
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
                        options.values.push(result.data.id);
                        options.webUrls.push(result.data.webUrl);
                    }
                    options.index++;
                    that.document.site(options, callback);
                });
            }
            else {
                if (options.values.length > 0) {
                    that.document.library({ siteId: options.values.shift(), siteUrl: options.webUrls.shift() }, callback);
                }
                else {
                    that.document.error({ title: "Get site url failed." });
                }
            }
        },
        library: function (options, callback) {
            that.service.libraries(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _l = "";
                    $.each(result.data.value, function (i, d) {
                        if (decodeURI(that.filePath).toUpperCase().indexOf(decodeURI(d.webUrl).toUpperCase()) > -1) {
                            _l = d.name;
                            return false;
                        }
                    });
                    if (_l != "") {
                        that.document.file({ siteId: options.siteId, siteUrl: options.siteUrl, listName: _l, fileName: that.utility.fileName(that.filePath) }, callback);
                    }
                    else {
                        that.document.error({ title: "Get library name failed." });
                    }
                }
                else {
                    that.document.error({ title: "Get library name failed." });
                }
            });
        },
        file: function (options, callback) {
            that.service.item(options, function (result) {
                if (result.status == app.status.succeeded) {
                    var _d = "";
                    $.each(result.data.value, function (i, d) {
                        if (decodeURI(that.filePath).toUpperCase() == decodeURI(d.EncodedAbsUrl).toUpperCase() && d.OData__dlc_DocId) {
                            _d = d.OData__dlc_DocId;
                            return false;
                        }
                    });
                    if (_d != "") {
                        that.documentId = _d;
                        that.popup.processing(false);
                        callback();
                    }
                    else {
                        that.document.error({ title: "Get file Document ID failed." });
                    }
                }
                else {
                    that.document.error({ title: "Get file Document ID failed." });
                }
            });
        },
        error: function (options) {
            that.controls.documentIdError.html("Error message: " + options.title);
            that.controls.main.addClass("error");
            that.popup.processing(false);
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
                that.controls.popupMain.removeClass("process confirm browse");
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
                    that.controls.popupMain.removeClass("process confirm browse").addClass("active message");
                }
                else {
                    that.controls.popupMessage.removeClass("error").addClass("success");
                    // that.controls.popupSuccessMessage.html(options.title);
                    that.controls.innerMessageBox.removeClass("active ms-MessageBar--success").addClass("active ms-MessageBar--error");
                    that.controls.innerMessageIcon.removeClass("ms-Icon--Completed").addClass("ms-Icon--ErrorBadge");
                    that.controls.innerMessageText.html(options.title);
                    $(".popups .bg").hide();
                    that.controls.popupMain.removeClass("process confirm browse");
                }
            }
            if (options.canClose) {
                that.controls.innerMessageBox.addClass("canclose");
            }
            else {
                that.controls.innerMessageBox.removeClass("canclose");
            }
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
            $(".popups .bg").removeAttr("style");
            if (!show) {
                that.controls.popupMain.removeClass("active process");
            }
            else {
                that.controls.popupMain.removeClass("message confirm browse").addClass("active process");
            }
        },
        confirm: function (options, yesCallback, noCallback) {
            $(".popups .bg").removeAttr("style");
            that.controls.popupConfirmTitle.html(options.title);
            that.controls.popupMain.removeClass("message process browse").addClass("active confirm");
            that.controls.popupConfirmYes.unbind("click").click(function () {
                yesCallback();
            });
            that.controls.popupConfirmNo.unbind("click").click(function () {
                noCallback();
            });
        },
        browse: function (show) {
            $(".popups .bg").removeAttr("style");
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
                    that.controls.innerMessageBox.removeClass("active");
                    if (!that.controls.popupMain.hasClass("confirm") && !that.controls.popupMain.hasClass("browse") && !that.controls.popupMain.hasClass("process")) {
                        that.controls.popupMain.removeClass("active message");
                    }
                }, millisecond);
            } else {
                $(".popups .bg").removeAttr("style");
                that.controls.innerMessageBox.removeClass("active");
                if (!that.controls.popupMain.hasClass("confirm") && !that.controls.popupMain.hasClass("browse") && !that.controls.popupMain.hasClass("process")) {
                    that.controls.popupMain.removeClass("active message");
                }
            }
        },
        back: function (millisecond) {
            if (millisecond) {
                setTimeout(function () {
                    $(".popups .bg").removeAttr("style");
                    that.controls.popupMain.removeClass("active message");
                    that.controls.main.removeClass("manage add edit bulk single table chart clone step-first step-second step-third step-fourth").addClass("manage");
                }, millisecond);
            }
            else {
                $(".popups .bg").removeAttr("style");
                that.controls.popupMain.removeClass("active message");
                that.controls.main.removeClass("manage add edit bulk single table chart clone step-first step-second step-third step-fourth").addClass("manage");
            }
        }
    };

    that.range = {
        create: function (options, callback) {
            Excel.run(function (ctx) {
                var p = that.utility.position(options.Position), r = ctx.workbook.worksheets.getItem(p.sheet).getRange(p.cell);
                ctx.workbook.bindings.add(r, Excel.BindingType.range, options.RangeId);
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        select: function (callback) {
            Excel.run(function (ctx) {
                var r = ctx.workbook.getSelectedRange();
                r.load("address");
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded, data: r.address });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: "Get selected cell failed." });
            });
        },
        exist: function (options, callback) {
            var _sourceType = options.Data.SourceType;
            if ((_sourceType == app.sourceTypes.table || _sourceType == app.sourceTypes.chart) && options.IsNameRange) {
                callback({ status: app.status.succeeded, data: { text: options.Data.Name, address: null } });
            }
            else {
                var _rid = options.IsNameRange ? options.Data.NameRangeId : options.Data.RangeId;
                if (_sourceType == app.sourceTypes.point || _sourceType == app.sourceTypes.table) {
                    Excel.run(function (ctx) {
                        var r = ctx.workbook.bindings.getItem(_rid).getRange();
                        r.load("text,address");
                        return ctx.sync().then(function () {
                            if (_sourceType == app.sourceTypes.point) {
                                callback({ status: app.status.succeeded, data: { text: $.trim(r.text[0][0]), address: r.address } });
                            }
                            else {
                                var _image = r.getImage();
                                return ctx.sync().then(function () {
                                    var _imageSrc = _image.value;
                                    that.range.formats({ range: r }, function (result) {
                                        if (result.status == app.status.succeeded) {
                                            callback({ status: app.status.succeeded, data: { text: JSON.stringify({ image: _imageSrc, table: result.data }), address: r.address } });
                                        }
                                        else {
                                            callback({ status: app.status.failed, message: "Get source table styles failed." });
                                        }
                                    });
                                });
                            }
                        });
                    }).catch(function (error) {
                        callback({ status: app.status.failed, message: error.message });
                    });
                }
                else {
                    Excel.run(function (ctx) {
                        var _p = that.utility.position(options.Data.Position);
                        var _chart = ctx.workbook.worksheets.getItem(_p.sheet).charts.getItem(_p.cell);
                        var _image = _chart.getImage();
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded, data: { text: _image.value, address: options.Data.Position } });
                        });
                    }).catch(function (error) {
                        callback({ status: app.status.failed, message: error.message });
                    });
                }
            }
        },
        goto: function (options, callback) {
            Excel.run(function (ctx) {
                var r = ctx.workbook.bindings.getItem(options.RangeId).getRange();
                r.select();
                r.load("text,address");
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded, data: { address: r.address } });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        gotoWorkSheet: function (options, callback) {
            Excel.run(function (ctx) {
                var w = ctx.workbook.worksheets.getItem(options.Sheet);
                w.activate();
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        del: function (options, callback) {
            Excel.run(function (ctx) {
                ctx.workbook.bindings.getItem(options.RangeId).delete();
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        workSheets: function (callback) {
            Excel.run(function (ctx) {
                var worksheets = ctx.workbook.worksheets, _s = [];
                worksheets.load('items');
                return ctx.sync().then(function () {
                    for (var i = 0; i < worksheets.items.length; i++) {
                        _s.push(worksheets.items[i].name);
                    }
                    callback({ status: app.status.succeeded, data: _s });
                });
            }).catch(function (error) {
                callback({ status: app.status.succeeded, message: "Load worksheets failed." });
            });
        },
        charts: function (options, callback) {
            if (typeof (options.index) === "undefined") {
                options.index = 0;
                options.charts = [];
            }
            if (options.index < options.sheets.length) {
                Excel.run(function (ctx) {
                    var charts = ctx.workbook.worksheets.getItem(options.sheets[options.index]).charts;
                    charts.load('items');
                    return ctx.sync().then(function () {
                        for (var i = 0; i < charts.items.length; i++) {
                            options.charts.push({ sheet: options.sheets[options.index], chart: charts.items[i].name });
                        }
                        options.index++;
                        that.range.charts(options, callback);
                    });
                }).catch(function (error) {
                    callback({ status: app.status.failed, message: "Load charts in worksheet failed." });
                });
            }
            else {
                callback({ status: app.status.succeeded, data: options.charts });
            }
        },
        formats: function (options, callback) {
            if (typeof (options.index) == "undefined") {
                var _position = that.utility.position(options.range.address);
                options.index = 0;
                options.formats = [];
                options.sheet = _position.sheet;
                options.address = _position.cell;
                options.cells = [];
                $.each(options.range.text, function (m, n) {
                    $.each(n, function (x, y) {
                        options.cells.push({ row: m, column: x });
                    });
                });
            }
            if (options.index < options.cells.length) {
                Excel.run(function (ctx) {
                    var range = ctx.workbook.worksheets.getItem(options.sheet).getRange(options.address);
                    var cell = range.getCell(options.cells[options.index].row, options.cells[options.index].column);
                    cell.load(["format/*", "format/borders", "format/fill", "format/font"]);
                    return ctx.sync().then(function () {
                        var format = {
                            columnWidth: cell.format.columnWidth,
                            horizontalAlignment: cell.format.horizontalAlignment,
                            rowHeight: cell.format.rowHeight,
                            verticalAlignment: cell.format.verticalAlignment,
                            wrapText: cell.format.wrapText,
                            fill: { color: cell.format.fill.color },
                            font: { bold: cell.format.font.bold, color: cell.format.font.color, italic: cell.format.font.italic, name: cell.format.font.name, size: cell.format.font.size, underline: cell.format.font.underline },
                            borders: []
                        };

                        var borders = cell.format.borders;
                        borders.load("items");
                        return ctx.sync().then(function () {
                            for (var i = 0; i < borders.items.length; i++) {
                                format.borders.push({ color: borders.items[i].color, sideIndex: borders.items[i].sideIndex, style: borders.items[i].style, weight: borders.items[i].weight });
                            }
                            format.borders.sort(function (_m, _n) {
                                return _m.sideIndex > _n.sideIndex ? 1 : _m.sideIndex < _n.sideIndex ? -1 : 0;
                            });
                            options.formats.push(format);
                            options.index++;
                            that.range.formats(options, callback);
                        });
                    });
                }).catch(function (error) {
                    callback({ status: app.status.failed });
                });
            }
            else {
                callback({ status: app.status.succeeded, data: { values: options.range.text, formats: app.formats(options.formats) } });
            }
        },
        values: function (options, callback) {
            if (options.refresh) {
                if (typeof (options.index) == "undefined") {
                    options.index = 0;
                }
                if (options.index < options.data.length) {
                    var _item = options.data[options.index], _sourceType = _item.SourceType;
                    that.range.exist({ Data: _item, IsNameRange: true }, function (ret) {
                        that.range.exist({ Data: _item, IsNameRange: false }, function (result) {
                            var _s = result.status == app.status.succeeded,
                                _st = ret.status == app.status.succeeded;
                            var _v = _s ? result.data.text : "",
                                _n = _st ? (_sourceType == app.sourceTypes.point ? ret.data.text : _item.Name) : "",
                                _pv = _item.PublishedHistories && _item.PublishedHistories.length > 0 ? (_item.PublishedHistories[0].Value ? _item.PublishedHistories[0].Value : "") : ""
                            _p = _s ? result.data.address : "",
                            _np = _st ? ret.data.address : "";
                            options.data[options.index].Value = _v;
                            options.data[options.index].Position = _p;
                            options.data[options.index].Name = _n;
                            options.data[options.index].NamePosition = _np;
                            options.data[options.index].PublishedStatus = _v == _pv;
                            options.data[options.index].NameExisted = _st;
                            options.data[options.index].ValueExisted = _s;
                            options.data[options.index].Existed = _s && _st;
                            options.index++;
                            that.range.values(options, callback);
                        });
                    });
                }
                else {
                    callback({ data: options.data });
                }
            }
            else {
                callback({ data: options.data });
            }
        }
    };

    that.service = {
		common: function (options, callback)
		{
			let apiToken = that.token;
			let apiHeaders = { "authorization": "Bearer " + apiToken };

            $.ajax({
                url: options.url,
                type: options.type,
                cache: false,
                data: options.data ? options.data : "",
                dataType: options.dataType,
				headers: options.headers ? options.headers : apiHeaders,
                success: function (data) {
                    callback({ status: app.status.succeeded, data: data });
                },
                error: function (error) {
                    if (error.status == 410) {
                        that.popup.message({ success: false, title: "The current login gets expired and needs re-authenticate. You will be redirected to the login page by click OK." }, function () {
                            //window.location = "";
                        });
                    }
                    else if (error.status == 202) {
                        callback({ status: app.status.succeeded });
                    }
                    else {
                        callback({ status: app.status.failed, error: error });
                    }
                }
            });
        },
        add: function (options, callback) {
            that.service.common({ url: that.endpoints.add, type: "POST", data: options.data, dataType: "json" }, callback);
        },
        edit: function (options, callback) {
            that.service.common({ url: that.endpoints.edit, type: "PUT", data: options.data, dataType: "json" }, callback);
        },
        list: function (callback) {
            that.service.common({ url: that.endpoints.list + "fileName=" + that.filePath + "&documentId=" + that.documentId, type: "GET", dataType: "json" }, callback);
        },
        del: function (options, callback) {
            that.service.common({ url: that.endpoints.del + options.Id, type: "DELETE" }, callback);
        },
        publish: function (options, callback) {
            that.service.common({ url: that.endpoints.publish, type: "POST", data: options.data, dataType: "json" }, callback);
        },
        associated: function (options, callback) {
            that.service.common({ url: that.endpoints.associated + options.Id, type: "GET", dataType: "json" }, callback);
        },
        deleteSelected: function (options, callback) {
            that.service.common({ url: that.endpoints.deleteSelected, type: "POST", data: options.data }, callback);
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
        filesInFolder: function (options, callback) {
            var _c = "<View Scope='RecursiveAll'><Query><Where><BeginsWith><FieldRef Name='EncodedAbsUrl' /><Value Type='Text'>" + (options.url + "/") + "</Value></BeginsWith></Where></Query></View>";
            var _d = JSON.stringify({ "query": { "__metadata": { "type": "SP.CamlQuery" }, "ViewXml": _c } });
            that.service.common({ url: options.siteUrl + "/_api/web/lists/getbytitle('" + options.listName + "')/GetItems?$expand=Folder,File&$select=FileLeafRef,EncodedAbsUrl,OData__dlc_DocId,FileSystemObjectType,Folder/ServerRelativeUrl", type: "POST", dataType: "json", data: _d, headers: { "authorization": "Bearer " + that.api.sharePointToken, "Content-type": "application/json; odata=verbose" } }, callback);
        },
        checkCloneStatus: function (options, callback) {
            that.service.common({ url: that.endpoints.checkCloneStatus, type: "POST", data: options.data }, callback);
        },
        copyFile: function (options, callback) {
            var _d = JSON.stringify({ "parentReference": { "driveId": that.action.clone.destination.listId, "id": that.action.clone.destination.id }, "name": options.Name });
            that.service.common({ url: that.endpoints.graph + "/sites/" + that.action.clone.source.siteId + "/drives/" + that.action.clone.source.listId + "/items/" + options.Id + "/Copy", type: "POST", dataType: "json", data: _d, headers: { "authorization": "Bearer " + that.api.token, "Content-type": "application/json" } }, callback);
        },
        getFile: function (options, callback) {
            var _f = (decodeURI(options.Url).toLocaleLowerCase()).replace((decodeURI(that.action.clone.source.siteUrl).toLocaleLowerCase() + "/"), ""), _fp = _f.substr(_f.indexOf("/") + 1);
            that.service.common({ url: that.endpoints.graph + "/sites/" + that.action.clone.source.siteId + "/drives/" + that.action.clone.source.listId + "/root:/" + _fp, type: "GET", dataType: "json", headers: { "authorization": "Bearer " + that.api.token } }, callback);
        },
        clone: function (options, callback) {
            that.service.common({ url: that.endpoints.cloneFiles, type: "POST", data: options.data }, callback);
        },
        newFolder: function (options, callback) {
            var _u = options.itemId != null ? that.endpoints.graph + "/sites/" + options.siteId + "/drives/" + options.listId + "/items/" + options.itemId + "/children" :
                that.endpoints.graph + "/sites/" + options.siteId + "/drives/" + options.listId + "/root/children";
            var _j = JSON.stringify({ "name": options.name, "folder": {} });
            that.service.common({ url: _u, type: "POST", data: _j, dataType: "json", headers: { "authorization": "Bearer " + that.api.token, "Content-type": "application/json" } }, callback);
        },
        userInfo: function (callback) {
            that.service.common({ url: that.endpoints.userInfo, type: "GET" }, callback);
        }
    };

    that.errorFilter = {
        errorFilterClass: '.error-point-filter',
        $errorFilter: function () {
            return $(that.errorFilter.errorFilterClass);
        },
        isErrorFiltered: function () {
            return that.errorFilter.$errorFilter().find('input').prop('checked');
        },
        doErrorFilter: function ($el, event) {
            //ensure click checkbox
            if (event.target.nodeName.toLowerCase() !== 'input') {
                $el.find('input').get(0).click();
            } else {
                that.utility.pager.init({ refresh: false });
            }
        },
        filterListByError: function (original) {
            if (!original || original.length === 0)
                return original;
            if (that.errorFilter.isErrorFiltered() === true)
                return original.filter(function (item) { return item.Existed !== true; });
            return original;
        },
        init: function () {
            that.errorFilter.$errorFilter().click(function (event) {
                that.errorFilter.doErrorFilter($(this), event);
            });
        }
    };

    that.ui = {
        associated: function (options) {
            var _s = [], _st = [];
            $.each(options.data, function (i, d) {
                if ($.inArray(d.Catalog.Name, _st) == -1) {
                    _s.push(that.utility.fileName(d.Catalog.Name));
                    _st.push(d.Catalog.Name);
                }
            });
            _s.sort(function (_a, _b) {
                return (_a.toUpperCase() > _b.toUpperCase()) ? 1 : (_a.toUpperCase() < _b.toUpperCase()) ? -1 : 0;
            });
            $.each(_s, function (i, d) {
                $("<li>" + d + "</li>").appendTo(that.controls.listAssociated);
            });
        },
        listFilter: {
            search: function (original) {
                var result = [];
                if (that.sourcePointKeyword != "") {
                    var _sk = app.search.splitKeyword({ keyword: that.sourcePointKeyword });
                    if (_sk.length > 26) {
                        that.popup.message({ success: false, title: "Only support less then 26 keywords." });
                    }
                    else {
                        $.each(original, function (i, d) {
                            if (app.search.weight({ keyword: _sk, source: d }) > 0) {
                                result.push(d);
                            }
                        });
                    }
                } else {
                    result = original;
                }
                return result;
            },
            filterError: function (original) {
                return that.errorFilter.filterListByError(original);
            }
        },
        list: function (options, callback) {
            that.range.values({ data: $.extend([], that.points), refresh: options.refresh }, function (result) {
                try {
                    var _dt = result.data, _d = [], _ss = [], _st = that.utility.sortType();
                    var _pt = that.utility.selectedSourceType(".point-types-mana");
                    // records that are filtered by source type and searching keyword 
                    var _df = [];
                    // counter for each source type
                    var _ia = 0, _ib = 0, _ic = 0;
                    var listFilter = that.ui.listFilter;
                    //search
                    _d = listFilter.search(_dt);
                    //error filter
                    _d = listFilter.filterError(_d);

                    // filtered by source type
                    if (_pt == app.sourceTypes.all) {
                        _df = _d;
                    }
                    else {
                        $.each(_d, function (_a, _b) {
                            if (_b.SourceType == _pt) {
                                _df.push(_b);
                            }
                        });
                    }

                    $.each(_d, function (i, d) {
                        if (d.SourceType == app.sourceTypes.point) {
                            _ia++;
                        } else if (d.SourceType == app.sourceTypes.chart) {
                            _ib++;
                        } else {
                            _ic++;
                        }
                    });
                    that.controls.sourceTypeNavMana[0].children[1].innerText = _ia;
                    that.controls.sourceTypeNavMana[1].children[1].innerText = _ib;
                    that.controls.sourceTypeNavMana[2].children[1].innerText = _ic;

                    that.utility.pager.status({ length: _df.length });
                    if (_st.sortType == app.sortTypes.name) {
                        if (_st.sortOrder == app.sortOrder.asc) {
                            _df.sort(function (_a, _b) {
                                return (app.string(_a.Name).toUpperCase() > app.string(_b.Name).toUpperCase()) ? 1 : (app.string(_a.Name).toUpperCase() < app.string(_b.Name).toUpperCase()) ? -1 : 0;
                            });
                        }
                        else {
                            _df.sort(function (_a, _b) {
                                return (app.string(_a.Name).toUpperCase() < app.string(_b.Name).toUpperCase()) ? 1 : (app.string(_a.Name).toUpperCase() > app.string(_b.Name).toUpperCase()) ? -1 : 0;
                            });
                        }
                    }
                    else if (_st.sortType == app.sortTypes.status) {
                        if (_st.sortOrder == app.sortOrder.asc) {
                            _df.sort(function (_a, _b) {
                                return (_a.PublishedStatus > _b.PublishedStatus) ? 1 : (_a.PublishedStatus < _b.PublishedStatus) ? -1 : 0;
                            });
                        }
                        else {
                            _df.sort(function (_a, _b) {
                                return (_a.PublishedStatus < _b.PublishedStatus) ? 1 : (_a.PublishedStatus > _b.PublishedStatus) ? -1 : 0;
                            });
                        }
                    }
                    else if (_st.sortType == app.sortTypes.value) {
                        if (_st.sortOrder == app.sortOrder.asc) {
                            _df.sort(function (_a, _b) {
                                return (_a.Value > _b.Value) ? 1 : (_a.Value < _b.Value) ? -1 : 0;
                            });
                        }
                        else {
                            _df.sort(function (_a, _b) {
                                return (_a.Value < _b.Value) ? 1 : (_a.Value > _b.Value) ? -1 : 0;
                            });
                        }
                    }
                    else {
                        _df.sort(function (_a, _b) {
                            return (app.string(_a.Name).toUpperCase() > app.string(_b.Name).toUpperCase()) ? 1 : (app.string(_a.Name).toUpperCase() < app.string(_b.Name).toUpperCase()) ? -1 : 0;
                        });
                    }

                    // get selected source points
                    var _c = that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked"), _s = that.utility.selected();
                    $.each(_s, function (m, n) {
                        _ss.push(n.SourcePointId);
                    });
                    that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", _c);
                    if (_c) {
                        that.controls.headerListPoints.find(".point-header .ckb-wrapper").addClass("checked");
                    }
                    else {
                        that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
                    }

                    that.utility.scrollTop();
                    that.controls.list.find(".point-item").remove();
                    that.ui.item({ index: 0, data: _df, selected: _ss }, callback);
                    if (that.controls.list.find(">li").length > 0) {
                        that.controls.moveUp.removeClass("disabled");
                        that.controls.moveDown.removeClass("disabled");
                    }
                } catch (err) {
                    that.popup.message({
                        success: false, title: "Error occurred: " + err.message
                    });
                }
            });
        },
        item: function (options, callback) {
            if (options.index < options.data.length) {
                var _item = options.data[options.index], _sourceType = _item.SourceType;

                if (options.index >= that.pagerSize * (that.pagerIndex - 1) && options.index < that.pagerSize * that.pagerIndex) {
                    var _pn = _item.NameExisted ? that.utility.position(_item.NamePosition) : {},
                        _pnt = _item.ValueExisted ? that.utility.position(_item.Position) : {},
                        _sel = $.inArray(_item.Id, options.selected) > -1,
                        _pht = that.utility.publishHistory({ data: _item.PublishedHistories }),
                        _cs = !_item.PublishedStatus ? "status-notpublished" : "status-published",
                        _ts = _sourceType == app.sourceTypes.point ? "type-point" : (_sourceType == app.sourceTypes.table ? "type-table" : "type-chart"),
                        _ttp = app.sourceTitle(_sourceType);
                    var _ta = false;
                    if (that.controls.sourceTypeNavMana.hasClass("is-selected")) {
                        _cs = _cs + " only";
                    }
                    else {
                        _ta = true;
                    }
                    var _h = '<li class="point-item' + (_item.Existed ? "" : " item-error") + ' ' + _cs + ' ' + _ts + '" data-id="' + _item.Id + '" data-range="' + _item.RangeId + '" data-position="' + _item.Position + '" data-namerange="' + _item.NameRangeId + '" data-nameposition="' + _item.NamePosition + '" data-sourcetype="' + _item.SourceType + '">';
                    _h += '<div class="point-item-line">';
                    // i1
                    _h += '<div class="i1"><div class="ckb-wrapper' + (_sel ? " checked" : "") + '"><input type="checkbox" ' + (_sel ? 'checked="checked"' : '') + ' /><i class="ms-Icon ms-Icon--CheckMark"></i></div></div>';
                    // i2
                    _h += '<div class="i2"><span class="s-name" title="' + _item.Name + '">' + _item.Name + '</span>';
                    _h += '<div class="sp-file-pos">';
                    if (_sourceType == app.sourceTypes.point) {
                        if (_item.NameExisted) {
                            _h += '<span title="' + (_pnt.sheet ? _pnt.sheet : "") + ':[' + (_pnt.cell ? _pnt.cell : "") + ']">' + (_pnt.sheet ? _pnt.sheet : "") + ':[' + (_pnt.cell ? _pnt.cell : "") + ']</span>';
                        }
                    }
                    else {
                        if (_item.ValueExisted) {
                            _h += '<span title="' + (_pnt.sheet ? _pnt.sheet : "") + ':[' + (_pnt.cell ? _pnt.cell : "") + ']">' + (_pnt.sheet ? _pnt.sheet : "") + ':[' + (_pnt.cell ? _pnt.cell : "") + ']</span>';
                        }
                    }
                    _h += '</div>';
                    _h += '</div>';
                    // i3
                    _h += '<div class="i3"><i class="ms-Icon ms-Icon--Error"></i><i class="ms-Icon ms-Icon--Completed"></i></div>';
                    // i4
                    if (_sourceType == app.sourceTypes.point) {
                        _h += '<div class="i4">';

                        if (_item.ValueExisted) {
                            var _spv = _item.Value && _item.Value != null ? _item.Value : "";
                            _h += '<span title="' + _spv + '">' + _spv + '</span>';
                        }

                        _h += '</div>';
                    }
                    else if (_ta) {
                        _h += '<div class="i4"></div>';
                    }
                    // i5
                    _h += '<div class="i5">';
                    _h += '<div class="i-menu"><a href="javascript:"><span title="Action"><i class="ms-Icon ms-Icon--More"></i></span><span class="quick-menu"><span class="i-history" title="History"><i class="ms-Icon ms-Icon--Clock"></i><i>History</i></span><span class="i-edit" title="Edit Custom Format"><i class="ms-Icon ms-Icon--Edit"></i><i>Edit</i></span><span class="i-delete" title="Delete"><i class="ms-Icon ms-Icon--Cancel"></i><i>Delete</i></span></span></a></div>';
                    _h += '</div>';
                    _h += '<div class="clear"></div>';
                    _h += '</div>';
                    // History
                    _h += '<div class="item-history"><ul class="history-list">';
                    _h += '<li class="history-header"><div class="h1"><span>Name</span></div><div class="h2"><span>Date Modified</span></div><div class="h3"><span>Value</span></div></li>';
                    $.each(_pht, function (m, n) {
                        var __cv = (n.Value ? n.Value : ""), __pvr = _sourceType == app.sourceTypes.point ? __cv : (__cv == "Cloned" ? "Cloned" : "Published");
                        _h += '<li class="history-item"><div class="h1" title="' + n.PublishedUser + '">' + n.PublishedUser +
                            '</div><div class="h2" title="' + app.date(n.PublishedDate) + '">' + app.date(n.PublishedDate) +
                            '</div><div class="h3" title="' + __pvr + '">' + __pvr + '</div></li>';
                    });
                    _h += '</ul>';
                    _h += '</div>';

                    // Error
                    _h += '<div class="error-info">';
                    _h += '<div class="e1"><i class="ms-Icon ms-Icon--ErrorBadge"></i></div>';
                    _h += '<div class="e2">';
                    _h += '<p><strong>Error</strong>: The source ' + _ttp + ' is invalid. This could be caused by not saving the excel file after creating the source point. Please delete the source ' + _ttp + ' and recreate it.</p>';
                    _h += '</div>';
                    _h += '</div>';
                    _h += '<div class="clear"></div>';
                    _h += '</li>';
                    that.controls.list.append(_h);
                }
                options.index++;
                that.ui.item(options, callback);
            }
            else {
                that.filteredPoints = options.data;
                if (callback) {
                    callback();
                }
            }
        },
        remove: function (options) {
            that.controls.list.find("[data-id=" + options.Id + "]").remove();
        },
        publish: function (options, callback) {
            $.each(options.SourcePoints, function (i, d) {
                var _e = that.controls.list.find("[data-id=" + d.Id + "]"),
                    _sourceType = d.SourceType
                _pv = (d.PublishedHistories && d.PublishedHistories.length > 0 ? (d.PublishedHistories[0].Value ? d.PublishedHistories[0].Value : "") : ""),
                _pht = that.utility.publishHistory({ data: d.PublishedHistories });
                _e.find(".history-list").find(".history-item").remove();
                _e.removeClass("status-notpublished").addClass("status-published");
                $.each(_pht, function (m, n) {
                    var __cv = (n.Value ? n.Value : ""), __pvr = _sourceType == app.sourceTypes.point ? __cv : (__cv == "Cloned" ? "Cloned" : "Published");
                    _e.find(".history-list").append('<li class="history-item"><div class="h1" title="' + n.PublishedUser + '">' + n.PublishedUser + '</div><div class="h2" title="' + __pvr + '">' + __pvr + '</div><div class="h3" title="' + app.date(n.PublishedDate) + '">' + app.date(n.PublishedDate) + '</div></li>');
                });
                if (d.SourceType == app.sourceTypes.point) {
                    _e.find(".i4").html('<span title="' + _pv + '">' + _pv + '</span>');
                }
                that.controls.list.find(".ckb-wrapper input").prop("checked", false);
                that.controls.list.find(".ckb-wrapper").removeClass("checked");
            });

            callback();
        },
        clone: function (options) {
            var _c = [], _nc = [];
            $.each(options.data ? options.data : [], function (_i, _d) {
                if (_d.Clone) {
                    _c.push(_d);
                }
                else {
                    _nc.push(_d);
                }
            });
            that.controls.listWillClone.html("");
            that.controls.listWillNotClone.html("");
            if (_c.length > 0) {
                that.controls.cloneBtn.removeClass("disabled ms-Button--primary").addClass("ms-Button--primary");
                $.each(_c, function (_i, _d) {
                    that.controls.listWillClone.append('<li><i class="ms-Icon ' + (_d.Name.toLowerCase().indexOf("xlsx") >= 0 ? 'ms-Icon--ExcelLogo' : 'ms-Icon--WordLogo') + '"></i><span>' + _d.Name + '</span></li>');
                });
            }
            else {
                that.controls.cloneBtn.removeClass("ms-Button--primary").addClass("disabled");
                that.controls.listWillClone.html("No items");
            }
            if (_nc.length > 0) {
                $.each(_nc, function (_i, _d) {
                    that.controls.listWillNotClone.append('<li><i class="ms-Icon ' + (_d.Name.toLowerCase().indexOf("xlsx") >= 0 ? 'ms-Icon--ExcelLogo' : 'ms-Icon--WordLogo') + '"></i><span>' + _d.Name + '</span></li>');
                });
            }
            else {
                that.controls.listWillNotClone.html("No items");
            }
        },
        cloneResult: function (options) {
            that.controls.main.removeClass("step-third").addClass("step-fourth");
            that.controls.cloneNext.removeClass("ms-Button--primary").addClass("disabled");
            that.controls.cloneBtn.removeClass("ms-Button--primary").addClass("disabled");
            that.controls.cloneResult.removeClass("clone-success clone-error");
            var _ec = [];
            var _rc = [];
            $.each(options.data, function (_i, _d) {
                if (!_d.Clone) {
                    _ec.push(_d);
                }
                else {
                    _rc.push(_d);
                }
            });

            if (_ec.length > 0) {
                that.controls.listDoneFail.html("");
                that.controls.cloneResult.addClass("clone-error");
                // that.popup.message({ success: false, title: "Get files in destination folder failed." });
                $.each(_ec, function (_i, _d) {
                    that.controls.listDoneFail.append('<li><i class="ms-Icon ms-Icon--ErrorBadge"></i><i class="ms-Icon ' + (_d.Name.toLowerCase().indexOf("xlsx") >= 0 ? 'ms-Icon--ExcelLogo' : 'ms-Icon--WordLogo') + '"></i><span>' + _d.Name + '</span></li>');
                });
            }
            if (_rc.length > 0) {
                that.controls.listDoneSuccess.html("");
                // that.controls.linkCloned.prop("href", that.action.clone.destination.url);
                that.controls.cloneResult.addClass("clone-success");
                $.each(_rc, function (_i, _d) {
                    that.controls.listDoneSuccess.append('<li><i class="ms-Icon ms-Icon--Completed"></i><i class="ms-Icon ' + (_d.Name.toLowerCase().indexOf("xlsx") >= 0 ? 'ms-Icon--ExcelLogo' : 'ms-Icon--WordLogo') + '"></i><span>' + _d.Name + '</span></li>');
                });
            }
        },
        charts: function (options, callback) {
            that.controls.chartList.html("");
            $.each(options.charts, function (i, d) {
                $('<li><a href="javascript:" data-address="' + d.sheet + '!' + d.chart + '" data-sheet="' + d.sheet + '" data-chart="' + d.chart + '">' + d.sheet + " - " + d.chart + '</a></li>').appendTo(that.controls.chartList);
            });
            callback();
        }
    };

    return that;
})();