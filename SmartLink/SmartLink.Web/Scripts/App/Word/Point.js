$(function ()
{
    Office.initialize = function (reason) {
        var isInIframe = function ()
        {
            try 
            {
                return window.self !== window.top;
            }
            catch(e)
            {
                return true;
            }
        };

		$(document).ready(function () 
		{
            BigNumber.config({ EXPONENTIAL_AT: 1e+9 });

			if (isInIframe())
            {
                microsoftTeams.initialize();

                microsoftTeams.authentication.authenticate({
                    url: '/auth',
                    width: 600,
                    height: 535,
                    successCallback: function (result) {
						point.init(result.idToken);
						$("#dvLogin").hide();
						$("#word-addin").show();
						
                    },
                    failureCallback: function (err) {
                        console.log(err);
                    }
                });
            }
            else
			{
                var authenticationContext = new AuthenticationContext(config);

                // Check For & Handle Redirect From AAD After Login
                if (authenticationContext.isCallback(window.location.hash)) {
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
								console.log("AzureAD error:", error, errorDesc);
                                authenticationContext.acquireTokenRedirect(config.clientId, null, null);
                            }
                            else
                            {
                                point.init(token);
                                $("#dvLogin").hide();
                                $("#word-addin").show();
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
		token: "",
        filePath: "",
        documentId: "",
        controls: {},
        file: null,
        selected: null,
        model: null,
        points: [],
        keyword: "",
        sourcePointKeyword: "",
        highlightColor: "#66FF00",
        highlighted: false,
        firstLoad: true,
        pagerIndex: 0,
        pagerSize: 30,
        pagerCount: 0,
        totalPoints: 0,
        endpoints: {
            add: "/api/DestinationPoint",
            catalog: "/api/SourcePointCatalog?documentId=",
            list: "/api/DestinationPointCatalog?",
            del: "/api/DestinationPoint?id=",
            deleteSelected: "/api/DeleteSelectedDestinationPoint",
            token: "/api/GraphAccessToken",
            sharePointToken: "/api/SharePointAccessToken",
            graph: "https://graph.microsoft.com/v1.0",
            customFormat: "/api/CustomFormats",
            updateCustomFormat: "/api/UpdateDestinationPointCustomFormat",
            recentFiles: "/api/RecentFiles",
            addRecentFile: "/api/RecentFile",
            checkCloneStatus: "/api/CloneCheckFile",
            cloneFiles: "/api/CloneFiles",
            userInfo: "/api/userprofile"
        },
        api: {
            host: "",
            token: "",
            sharePointToken: ""
        },
        recentFiles: []
    }, that = point;

    that.init = function (accessToken) {
		that.token = accessToken;
		that.filePath = Office.context.document.url;
        that.controls = {
            body: $("body"),
            main: $(".main"),
            back: $(".n-back"),
            add: $(".n-add"),
            highlight: $(".n-highlight"),
            refresh: $(".n-refresh"),
            del: $(".n-delete"),
            next: $("#btnNext"),
            cancel: $("#btnCancel"),
            file: $("#txtFile"),
            fileTrigger: $("#btnOpenBrowse"),
            keyword: $("#txtKeyword"),
            search: $("#iSearch"),
            autoCompleteControl: $("#autoCompleteWrap"),
            resultList: $("#resultList"),
            resultNotFound: $("#resultNotFound"),
            selectedName: $("#selectedName"),
            stepFirstMain: $(".add-point-first"),
            stepSecondMain: $(".add-point-second"),
            sourcePointName: $("#txtSearchSourcePoint"),
            searchSourcePoint: $("#iSearchSourcePoint"),
            autoCompleteControl2: $("#autoCompleteWrap2"),
            list: $("#listPoints"),
            documentIdError: $("#lblDocumentIDError"),
            documentIdReload: $("#btnDocumentIDReload"),
            headerListPoints: $("#headerListPoints"),
            headerListPointsAdd: $("#headerListPointsAdd"),
            moveUp: $("#btnMoveUp"),
            moveDown: $("#btnMoveDown"),
            previewValue: $("#lbPreviewValue"),
            addCustomFormat: $("#addCustomFormat"),
            recentFilesDrop: $(".ms-Dropdown-title"),
            recentFiles: $("#recentFiles"),
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
            sourceTypeNav: $(".point-types li"),
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
            popupConfirmMessage: $("#confirmMessage"),
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

        that.highlighted = that.utility.highlight.get();
        that.controls.highlight.find("span").html(that.highlighted ? "Highlight Off" : "Highlight On");
        that.controls.highlight.prop("title", that.highlighted ? "Highlight Off" : "Highlight On");
        that.controls.body.click(function () {
            that.action.body();
        });
        that.controls.back.click(function () {
            that.action.back();
        });
        that.controls.add.click(function () {
            that.action.add();
        });
        that.controls.highlight.click(function () {
            that.action.highlightAll();
        });
        that.controls.refresh.click(function () {
            that.action.refresh();
        });
        that.controls.del.click(function () {
            that.action.deleteSelected();
        });
        that.controls.cancel.click(function () {
            that.action.cancel();
        });
        that.controls.resort.click(function () {
            that.action.resort();
        });
        that.controls.file.click(function () {
            that.browse.browseFile = true;
            that.action.clone.source = null;
            that.action.clone.destination = null;
            that.browse.init();
        });
        that.controls.fileTrigger.click(function () {
            that.browse.browseFile = true;
            that.action.clone.source = null;
            that.action.clone.destination = null;
            that.browse.init();
        });
        that.controls.keyword.focus(function () {
            that.action.dft(this, true);
        });
        that.controls.keyword.blur(function () {
            that.action.dft(this, false);
        });
        that.controls.keyword.keydown(function (e) {
            if (e.keyCode == 13) {
                if (that.controls.keyword.val() != "") {
                    $(".search-tooltips").hide();
                    that.action.search(true);
                }
            }
            else if (e.keyCode == 38 || e.keyCode == 40) {
                app.search.move({ result: that.controls.autoCompleteControl, target: that.controls.keyword, down: e.keyCode == 40 });
            }
        });
        that.controls.keyword.bind("input", function (e) {
            that.action.autoComplete();
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
        that.controls.search.click(function () {
            that.action.search(!that.controls.keyword.closest(".input-search").hasClass("searched"));
        });
        that.controls.resultList.on("click", ".point-item", function () {
            that.action.choose($(this));
        });
        that.controls.popupErrorOK.click(function () {
            that.action.ok();
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

        that.controls.moveUp.click(function () {
            that.action.up();
        });
        that.controls.moveDown.click(function () {
            that.action.down();
        });
        that.controls.list.on("click", "li .btnSelectFormat", function () {
            $(this).closest(".point-item").find(".listFormats").hasClass("active") ? $(this).closest(".point-item").find(".listFormats").removeClass("active") : $(this).closest(".point-item").find(".listFormats").addClass("active");
            return false;
        });
        that.controls.list.on("click", "li .iconSelectFormat", function () {
            $(this).closest(".point-item").find(".listFormats").hasClass("active") ? $(this).closest(".point-item").find(".listFormats").removeClass("active") : $(this).closest(".point-item").find(".listFormats").addClass("active");
            return false;
        });
        that.controls.list.on("click", "li .listFormats ul > li", function () {
            var _ck = $(this).hasClass("checked"), _sg = $(this).closest(".drp-radio").length > 0, _cn = $(this).data("name");
            if (_sg) {
                $(this).closest("ul").find("li").removeClass("checked")
            }
            _ck ? $(this).removeClass("checked") : $(this).addClass("checked");
            if (_cn == "ConvertToThousands" || _cn == "ConvertToMillions" || _cn == "ConvertToBillions" || _cn == "ConvertToHundreds") {
                $(this).closest(".listFormats").removeClass("convert1 convert2 convert3 convert4");
                $(this).closest(".listFormats").find(".drp-descriptor li.checked").removeClass("checked");
                if (!_ck) {
                    var _tn = _cn == "ConvertToThousands" ? "IncludeThousandDescriptor" : (_cn == "ConvertToMillions" ? "IncludeMillionDescriptor" : (_cn == "ConvertToBillions" ? "IncludeBillionDescriptor" : (_cn == "ConvertToHundreds" ? "IncludeHundredDescriptor" : "")));
                    var _cl = _cn == "ConvertToThousands" ? "convert2" : (_cn == "ConvertToMillions" ? "convert3" : (_cn == "ConvertToBillions" ? "convert4" : (_cn == "ConvertToHundreds" ? "convert1" : "")));
                    $(this).closest(".listFormats").addClass(_cl);
                }
            }
            that.action.selectedFormats($(this));
            return false;
        });
        that.controls.resultList.on("click", "li .btnSelectFormat", function () {
            $(this).closest(".point-item").find(".listFormats").hasClass("active") ? $(this).closest(".point-item").find(".listFormats").removeClass("active") : $(this).closest(".point-item").find(".listFormats").addClass("active");
            return false;
        });
        that.controls.resultList.on("click", "li .iconSelectFormat", function () {
            $(this).closest(".point-item").find(".listFormats").hasClass("active") ? $(this).closest(".point-item").find(".listFormats").removeClass("active") : $(this).closest(".point-item").find(".listFormats").addClass("active");
            return false;
        });
        that.controls.resultList.on("click", "li.selected .listFormats ul > li", function () {
            var _ck = $(this).hasClass("checked"), _sg = $(this).closest(".drp-radio").length > 0, _cn = $(this).data("name");
            if (_sg) {
                $(this).closest("ul").find("li").removeClass("checked")
            }
            _ck ? $(this).removeClass("checked") : $(this).addClass("checked");
            if (_cn == "ConvertToThousands" || _cn == "ConvertToMillions" || _cn == "ConvertToBillions" || _cn == "ConvertToHundreds") {
                that.controls.resultList.find("li.selected .listFormats").removeClass("convert1 convert2 convert3 convert4");
                that.controls.resultList.find("li.selected .listFormats .drp-descriptor li.checked").removeClass("checked");
                if (!_ck) {
                    var _tn = _cn == "ConvertToThousands" ? "IncludeThousandDescriptor" : (_cn == "ConvertToMillions" ? "IncludeMillionDescriptor" : (_cn == "ConvertToBillions" ? "IncludeBillionDescriptor" : (_cn == "ConvertToHundreds" ? "IncludeHundredDescriptor" : "")));
                    var _cl = _cn == "ConvertToThousands" ? "convert2" : (_cn == "ConvertToMillions" ? "convert3" : (_cn == "ConvertToBillions" ? "convert4" : (_cn == "ConvertToHundreds" ? "convert1" : "")));
                    that.controls.resultList.find("li.selected .listFormats").addClass(_cl);
                }
            }
            that.action.selectedFormats($(this));
            return false;
        });
        that.controls.resultList.on("click", "li .i-increase", function () {
            that.action.increase($(this));
            return false;
        });
        that.controls.resultList.on("click", "li .i-decrease", function () {
            that.action.decrease($(this));
            return false;
        });
        that.controls.resultList.on("click", "li .i-add", function () {
            $(this).blur();
            that.action.save($(this));
            return false;
        });
        that.controls.documentIdReload.click(function () {
            window.location.reload();
        });
        that.controls.recentFilesDrop.click(function () {
            $(this).closest(".recent-files").hasClass("is-open") ? $(this).closest(".recent-files").removeClass("is-open") : $(this).closest(".recent-files").addClass("is-open");
        });
        that.controls.recentFilesDrop.blur(function () {
            $(this).closest(".recent-files").removeClass("is-open");
        });
        that.controls.recentFiles.on("mousedown", "li", function (e) {
            $(this).parent().find("li.is-selected").removeClass("is-selected");
            $(this).addClass("is-selected");
            that.action.selectRecentFile({ elem: $(this) });
        });

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

            if ($(this).data("content") != "Points" && !_t) {
                that.controls.headerListPoints.find(".i3 span")[0].innerText = "Published Status";
            }
            else {
                that.controls.headerListPoints.find(".i3 span")[0].innerText = "Value";
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
            if ($(this).hasClass("is-selected")) {
                $(this).removeClass("is-selected");
            } else {
                that.controls.sourceTypeNav.removeClass("is-selected");
                $(this).addClass("is-selected");
            }

            if ($(this).data("content") != "Points") {
                that.controls.headerListPointsAdd.find(".i3").hide();
            }
            else {
                that.controls.headerListPointsAdd.find(".i3").show();
            }
            that.ui.sources({ data: that.file, keyword: that.keyword, sourceType: that.utility.selectedSourceType(".point-types") });
        });
        that.controls.headerListPoints.find(".i2,.i3,.i4").click(function () {
            that.action.sort($(this));
        });
        that.controls.headerListPointsAdd.find(".i2,.i3,.i4").click(function () {
            that.action.sortAdd($(this));
        });
        /* Source table and chart end */

        that.controls.list.on("click", ".i-edit", function () {
            that.action.edit($(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".i-update", function () {
            $(this).blur();
            that.action.update($(this));
            return false;
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
            that.action.del($(this).closest(".point-item"));
            return false;
        });
        that.controls.list.on("click", ".i-increase", function () {
            that.action.increase($(this));
            return false;
        });
        that.controls.list.on("click", ".i-decrease", function () {
            that.action.decrease($(this));
            return false;
        });
        that.controls.list.on("click", ".point-item", function (e) {
            that.action.goto($(this));
            if ($(this).closest(".point-item").hasClass("item-more")) {
                $(this).closest(".point-item").removeClass("item-more");
            }
        });
        that.controls.main.on("change", ".ckb-wrapper input", function (e) {
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
        that.controls.list.on("click", ".i-file", function () {
            that.action.open($(this));
            return false;
        });
        that.controls.main.on("click", ".search-tooltips li", function () {
            $(this).parent().parent().find("input").val($(this).text());
            $(this).parent().hide();
            if ($(this).closest(".input-search").length > 0) {
                that.action.search(true);
            }
            else {
                that.action.searchSourcePoint(true);
            }
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
        });
        that.utility.height();
        that.action.dft(that.controls.sourcePointName, false);

        // Retrieve the document ID via document URL
        that.document.init(function () {
            that.list({ refresh: false, index: 1 }, function (result) {
                if (result.status == app.status.succeeded) {
                    that.popup.processing(false);
                }
                else {
                    that.popup.message({ success: false, title: result.error.statusText });
                }
            });
            // Get user Info
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
                    that.points = result.data.DestinationPoints;
                    that.utility.pager.init({ refresh: options.refresh, index: options.index }, callback);
                }
                else {
                    that.utility.pager.status({ length: 0 });
                    callback({ status: app.status.succeeded });
                }
            }
            else {
                callback({ status: result.status, error: result.error });
            }
        });
    };

    that.default = function () {
        that.controls.main.removeClass("manage add edit clone step-first step-second step-third step-fourth").addClass("add");
        that.ui.reset();
    };

    that.files = function (options) {
        that.controls.recentFiles.find("li.is-selected").removeClass("is-selected");
        that.popup.processing(true);
        that.service.catalog({ documentId: options.documentId }, function (result) {
            if (result.status == app.status.succeeded) {
                that.popup.processing(false);
                that.file = result.data;
                that.ui.reset();
                that.action.select(options);
            }
            else {
                that.popup.message({ success: false, title: "Load source points failed." });
            }
        });
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
            callback();
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
            var file = $.trim(that.controls.file.val()), fd = that.controls.file.data("default"), keyword = $.trim(that.controls.keyword.val()), kd = that.controls.keyword.data("default");
            return { file: file != fd ? file : "", keyword: keyword != kd ? keyword : "" };
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
        update: function (options) {
            that.points[that.utility.index(options)] = options;
        },
        add: function (options) {
            that.points.push(options);
        },
        remove: function (options) {
            that.points.splice(that.utility.index(options), 1);
        },
        fileName: function (path) {
            return path.lastIndexOf("/") > -1 ? path.substr(path.lastIndexOf("/") + 1) : (path.lastIndexOf("\\") > -1 ? path.substr(path.lastIndexOf("\\") + 1) : path);
        },
        filePath: function (path, libraryPath) {
            return decodeURI(path).replace(decodeURI(libraryPath), "");
        },
        highlight: {
            get: function () {
                try {
                    var _d = Office.context.document.settings.get("HighLightPoint");
                    return _d && _d != null ? _d : false;
                } catch (e) {
                    return false;
                }
            },
            set: function (options, callback) {
                try {
                    Office.context.document.settings.set("HighLightPoint", options);
                    Office.context.document.settings.saveAsync({ overwriteIfStale: true }, function (asyncResult) {
                        callback({ status: asyncResult.status == Office.AsyncResultStatus.Succeeded ? app.status.succeeded : app.status.failed });
                    });
                } catch (e) {
                    callback({ status: app.status.failed });
                }
            }
        },
        pager: {
            init: function (options, callback) {
                that.controls.pagerValue.val("");
                that.controls.indexes.html("");
                that.pagerIndex = options.index ? options.index : 1;
                that.ui.list({ refresh: options.refresh }, callback);
                // that.utility.pager.updatePager();
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
        order: function (callback) {
            var _d = $.extend([], that.points);
            that.range.all(function (result) {
                if (result.status == app.status.succeeded) {
                    $.each(_d, function (i, d) {
                        var __i = that.range.index({ data: result.data, tag: d.RangeId });
                        _d[i].orderBy = __i > -1 ? __i : 99999;
                    });
                }
                else {
                    $.each(_d, function (i, d) {
                        _d[i].orderBy = 99999;
                    });
                }

                _d.sort(function (_a, _b) {
                    return (_a.orderBy > _b.orderBy) ? 1 : (_a.orderBy < _b.orderBy) ? -1 : 0;
                });

                callback({ status: app.status.succeeded, data: _d });
            });
        },
        selected: function () {
            var _s = [];
            that.controls.list.find(".point-item .ckb-wrapper input").each(function (i, d) {
                if ($(d).prop("checked")) {
                    var _id = $(d).closest(".point-item").data("id"), _rid = $(d).closest(".point-item").data("range");
                    _s.push({ DestinationPointId: _id, RangeId: _rid, Deleted: $(d).closest(".point-item").hasClass("item-error") });
                }
            });
            return _s;
        },
        height: function () {

        },
        unSelectAll: function () {
            that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", false);
            that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
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
        selectedSource: function (id) {
            var _m = null;
            if (id && that.file.SourcePoints) {
                var _t = [];
                $.each(that.file.SourcePoints, function (i, d) {
                    _t.push(d.Id);
                });
                if ($.inArray(id, _t) > -1) {
                    _m = that.file.SourcePoints[$.inArray(id, _t)];
                }
            }
            return _m;
        },
        sortType: function () {
            if (that.controls.headerListPoints.find(".i2.sort-desc,.i2.sort-asc").length > 0) {
                return { sortType: app.sortTypes.name, sortOrder: that.controls.headerListPoints.find(".i2").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
            else if (that.controls.headerListPoints.find(".i3.sort-desc,.i3.sort-asc").length > 0) {
                return { sortType: app.sortTypes.value, sortOrder: that.controls.headerListPoints.find(".i3").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
            else if (that.controls.headerListPoints.find(".i4.sort-desc,.i4.sort-asc").length > 0) {
                return { sortType: app.sortTypes.type, sortOrder: that.controls.headerListPoints.find(".i4").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
            else {
                return { sortType: app.sortTypes.df };
            }
        },
        sortTypeAdd: function () {
            if (that.controls.headerListPointsAdd.find(".i2.sort-desc,.i2.sort-asc").length > 0) {
                return { sortType: app.sortTypes.name, sortOrder: that.controls.headerListPointsAdd.find(".i2").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
            else if (that.controls.headerListPointsAdd.find(".i3.sort-desc,.i3.sort-asc").length > 0) {
                return { sortType: app.sortTypes.value, sortOrder: that.controls.headerListPointsAdd.find(".i3").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
            else if (that.controls.headerListPointsAdd.find(".i4.sort-desc,.i4.sort-asc").length > 0) {
                return { sortType: app.sortTypes.type, sortOrder: that.controls.headerListPointsAdd.find(".i4").hasClass("sort-desc") ? app.sortOrder.desc : app.sortOrder.asc };
            }
        }
    };

    that.action = {
        body: function () {
            $(".search-tooltips").hide();
            //that.controls.formatList.removeClass("active");
        },
        add: function () {
            that.utility.mode(function () {
                that.popup.processing(true);
                that.controls.resultList.find("li.selected").removeClass("selected");
                that.service.recentFiles(function (result) {
                    if (result.status == app.status.succeeded) {
                        that.popup.processing(false);
                        that.recentFiles = result.data;
                        that.ui.recentFiles({ data: result.data }, function () {
                            that.selected = null;
                            that.default();
                            if (that.controls.recentFiles.find("li").length > 0) {
                                that.action.selectRecentFile({ elem: that.controls.recentFiles.find("li").eq(0) });
                            }
                        });
                    }
                    else {
                        that.popup.message({ success: false, title: "Load recent files failed." });
                    }
                });
            });
        },
        back: function () {
            if (that.controls.main.hasClass("add")) {
                if (that.controls.main.hasClass("edit")) {
                    that.controls.main.removeClass("add edit clone step-first step-third step-fourth").addClass("manage");
                    that.utility.scrollTop();
                }
                else if (that.controls.main.hasClass("clone")) {
                    if (that.controls.main.hasClass("step-fourth")) {
                        //that.controls.cloneNext.addClass("disabled");
                        //that.controls.cloneBtn.addClass("disabled");
                        //that.controls.main.removeClass("step-fourth").addClass("step-third");
                        //that.action.clone.check();
                        // Do nothing
                        // that.controls.main.removeClass("add edit clone step-first step-third step-fourth").addClass("manage");
                        // that.utility.scrollTop();
                    }
                    else if (that.controls.main.hasClass("step-third")) {
                        that.controls.cloneNext.removeClass("disabled ms-Button--primary").addClass("ms-Button--primary");
                        that.controls.cloneBtn.removeClass("ms-Button--primary").addClass("disabled");
                        that.controls.main.removeClass("step-third").addClass("step-first");
                    }
                    else {
                        that.controls.main.removeClass("add edit clone step-first step-third step-fourth").addClass("manage");
                        that.utility.scrollTop();
                    }
                }
                else {
                    if (that.controls.main.hasClass("step-second")) {
                        that.ui.status({ next: true });
                        that.controls.main.removeClass("step-second").addClass("step-first");
                    }
                    else {
                        that.controls.main.removeClass("add edit clone step-first step-second step-third step-fourth").addClass("manage");
                        that.utility.scrollTop();
                    }
                }
            }
        },
        backToList: function () {
            that.controls.main.removeClass("add edit clone step-first step-second step-third step-fourth").addClass("manage");
            that.utility.scrollTop();
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
        select: function (options) {
            that.controls.file.val(options.name).removeClass("input-default");
            that.ui.select();
            that.selected = null;
            that.controls.sourceTypeNav.removeClass("is-selected");
            $(that.controls.recentFilesDrop[0]).text(options.name);
            that.ui.sources({ data: that.file, keyword: "", sourceType: that.utility.selectedSourceType(".point-types") });
            that.controls.stepFirstMain.addClass("selected-file");
            that.utility.height();
        },
        search: function (s) {
            var _e = that.utility.entered();
            if (s && _e.keyword != "") {
                that.controls.keyword.closest(".input-search").removeClass("searched").addClass("searched");
            }
            else {
                _e.keyword = "";
                that.controls.keyword.val("");
                that.action.dft(that.controls.keyword, false);
                that.controls.keyword.closest(".input-search").removeClass("searched");
            }
            that.keyword = _e.keyword;
            that.selected = null;
            that.controls.sourceTypeNav.removeClass("is-selected");
            if (_e.keyword != "") {
                that.controls.search.removeClass("ms-Icon--Search ms-Icon--Cancel").addClass("ms-Icon--Cancel");
            }
            else {
                _e.keyword = "";
                that.controls.keyword.val("");
                that.action.dft(that.controls.keyword, false);
                that.controls.search.removeClass("ms-Icon--Cancel").addClass("ms-Icon--Search");
            }
            that.ui.sources({ data: that.file, keyword: _e.keyword, sourceType: that.utility.selectedSourceType(".point-types") });
        },
        choose: function (o) {
            var _pointList = o.closest(".point-list");
            var _selectedItem = o.hasClass("point-item") ? o : o.closest(".point-item");
            var _i = _selectedItem.data("id"), _s = that.utility.selectedSource(_i);
            if (!that.selected || (that.selected && that.selected.Id != _i)) {
                that.selected = { Id: _i, File: _selectedItem.data("file"), Name: _selectedItem.data("name"), Value: _s.Value, SourceType: _s.SourceType };
                _pointList.find("li.selected .ckb-wrapper").removeClass("checked");
                _pointList.find("li.selected").removeClass("selected");
                _selectedItem.addClass("selected");
                _selectedItem.find(".ckb-wrapper").addClass("checked");
                _selectedItem.find(".lbPreviewValue").html(that.selected.Value);
                _selectedItem.find(".btnSelectFormat").prop("original", that.selected.Value);
                that.action.next({ o: _selectedItem, selectedPoint: that.selected });
            }

        },
        next: function (options) {
            if (that.selected.SourceType == app.sourceTypes.point) {
                that.action.customFormat(options);
            }
        },
        refresh: function () {
            that.list({ refresh: true, index: that.pagerIndex }, function (result) {
                if (result.status == app.status.failed) {
                    that.popup.message({ success: false, title: result.error.statusText });
                }
                else {
                    // that.controls.tooltipMessage.removeClass("active");
                    that.popup.message({ success: true, title: "Refresh all destination points succeeded." }, function () { that.popup.hide(3000); });
                }
            });
        },
        cancel: function () {
            if (that.controls.main.hasClass("edit")) {
                that.action.back();
            }
            else {
                that.ui.status({ next: true });
                that.controls.main.removeClass("step-third step-fourth").addClass("step-first");
            }
        },
        save: function (o) {
            o.closest(".point-item").find(".listFormats").removeClass("active");
            that.utility.mode(function () {
                var _s = o.closest(".point-item").find(".btnSelectFormat").prop("selected"), _n = o.closest(".point-item").find(".btnSelectFormat").prop("name"),
                    _f = (typeof (_n) != "undefined" && _n != "") ? _n.split(",") : [],
                    _c = (typeof (_s) != "undefined" && _s != "") ? _s.split(",") : [],
                    _fa = [],
                    _x = o.closest(".point-item").find(".btnSelectFormat").prop("place");
                $.each(_f, function (_a, _b) {
                    _fa.push({ Name: _b });
                });
                var _v = that.selected.SourceType == app.sourceTypes.point ? that.format.convert({ value: that.selected.Value, formats: _fa, decimal: _x }) : that.selected.Value;
                var _dpTypes = 0;
                if (that.selected.SourceType == app.sourceTypes.point) {
                    _dpTypes = app.destinationTypes.point;
                }
                else if (that.selected.SourceType == app.sourceTypes.table) {
                    if ($(o).hasClass("i-table-image")) {
                        _dpTypes = app.destinationTypes.tableImage;
                    }
                    else {
                        _dpTypes = app.destinationTypes.tableCell;
                    }
                }
                else {
                    _dpTypes = app.destinationTypes.chart;
                }
                var _json = $.extend({}, that.selected, { RangeId: app.guid(), CatalogName: that.filePath, CustomFormatIds: _c, Value: _v, DecimalPlace: _x, DestinationType: _dpTypes });

                that.range.create(_json, function (ret) {
                    if (ret.status == app.status.succeeded) {
                        that.popup.processing(true);
                        that.service.add({ data: { CatalogName: _json.CatalogName, DocumentId: that.documentId, RangeId: _json.RangeId, SourcePointId: _json.Id, CustomFormatIds: _json.CustomFormatIds, DecimalPlace: _json.DecimalPlace, DestinationType: _dpTypes } }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.range.all(function (ret) {
                                    var _index = -1;
                                    if (ret.status == app.status.succeeded) {
                                        _index = that.range.index({ data: ret.data, tag: _json.RangeId });
                                    }
                                    if (_index == -1) {
                                        _index = 99999;
                                    }
                                    that.utility.add($.extend({}, result.data, { existed: true, changed: false, orderBy: _index, PublishedStatus: true, DocumentValue: _v }));
                                    that.utility.pager.init({ refresh: false, index: that.pagerIndex }, function () {
                                        that.popup.message({ success: true, title: "Add new destination point succeeded." }, function () { that.popup.hide(3000); });
                                    });
                                });
                            }
                            else {
                                that.popup.message({ success: false, title: "Add destination point failed." });
                            }
                        });
                    }
                    else {
                        that.popup.message({ success: false, title: "Create range in Word failed." });
                    }
                });
            });
        },
        update: function (o) {
            that.action.backToList();
            o.closest(".point-item").find(".listFormats").removeClass("active");
            that.utility.mode(function () {
                var _s = o.closest(".point-item").find(".btnSelectFormat").prop("selected"), _n = o.closest(".point-item").find(".btnSelectFormat").prop("name"), _o = o.closest(".point-item").find(".btnSelectFormat").prop("original"),
                    _f = (typeof (_n) != "undefined" && _n != "") ? _n.split(",") : [],
                    _c = (typeof (_s) != "undefined" && _s != "") ? _s.split(",") : [],
                    _fa = [],
                    _x = o.closest(".point-item").find(".btnSelectFormat").prop("place");
                $.each(_f, function (_a, _b) {
                    _fa.push({ Name: _b });
                });
                var _v = that.format.convert({ value: _o, formats: _fa, decimal: _x });
                var _json = $.extend({}, {}, { Id: that.model.Id, RangeId: that.model.RangeId, CustomFormatIds: _c, Value: _v, DecimalPlace: _x, SourceType: that.model.ReferencedSourcePoint.SourceType, DestinationType: that.model.DestinationType });

                that.range.edit(_json, function (ret) {
                    if (ret.status == app.status.succeeded) {
                        that.popup.processing(true);
                        that.service.update({ data: { Id: _json.Id, CustomFormatIds: _json.CustomFormatIds, DecimalPlace: _json.DecimalPlace } }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.utility.update($.extend({}, result.data, { existed: true, changed: that.model.changed, orderBy: that.model.orderBy, PublishedStatus: that.model.PublishedStatus, DocumentValue: _v }));
                                that.utility.pager.init({ refresh: false, index: that.pagerIndex }, function () {
                                    that.popup.message({ success: true, title: "Update destination point custom format succeeded." }, function () { that.popup.hide(3000); });
                                });
                            }
                            else {
                                that.popup.message({ success: false, title: "Update destination point custom format failed." });
                            }
                        });
                    }
                    else {
                        that.popup.message({ success: false, title: "Update range in Word failed." });
                    }
                });
            });
        },
        highlightAll: function () {
            that.popup.processing(true);
            that.utility.order(function (result) {
                options = { index: 0, data: result.data, errorAmount: 0, successAmount: 0 };
                that.action.highlight(options);
            });
        },
        highlight: function (options) {
            if (options.index < options.data.length) {
                that.range.highlight(options.data[options.index], function (result) {
                    if (result.status == app.status.succeeded) {
                        options.successAmount++;
                    }
                    else {
                        options.errorAmount++;
                    }
                    options.index++;
                    that.action.highlight(options);
                });
            }
            else {
                that.highlighted = !that.highlighted;
                that.utility.highlight.set(that.highlighted, function () {
                    that.controls.highlight.find("span").html(that.highlighted ? "Highlight Off" : "Highlight On");
                    that.controls.highlight.prop("title", that.highlighted ? "Highlight Off" : "Highlight On");
                    that.popup.processing(false);
                });
            }
        },
        del: function (o) {
            that.utility.mode(function () {
                var _i = o.data("id"), _rid = o.data("range"), _dl = o.hasClass("item-error");
                that.popup.confirm({ title: "Do you want to delete the destination point?", message: "Removing this destination point cannot be undone." }, function () {
                    // Ask if user want to stay value in document
                    that.action.requestValueInDoc(_dl, function (keepContent) {

                        that.popup.processing(true);
                        that.service.del({ Id: _i }, function (result) {
                            if (result.status == app.status.succeeded) {
                                that.range.del({ RangeId: _rid }, function (result) {
                                    if (result.status == app.status.succeeded) {
                                        that.popup.message({ success: true, title: "Delete destination point succeeded." }, function () { that.popup.hide(3000); });
                                        that.utility.remove({ Id: _i });
                                        that.ui.remove({ Id: _i });
                                        that.utility.pager.init({ refresh: false, index: that.controls.list.find(".point-item").length > 0 ? that.pagerIndex : that.pagerIndex - 1 });
                                    }
                                    else {
                                        that.popup.message({ success: false, title: "Delete destination point in Word failed." });
                                    }
                                }, keepContent); // Add the parameter here to stay value or not.
                            }
                            else {
                                that.popup.message({ success: false, title: "Delete destination point failed." });
                            }
                        });

                    });

                }, function () {
                    that.controls.popupMain.removeClass("message process confirm active");
                });
            });
        },
        deleteSelected: function () {
            var _s = that.utility.selected(), _ss = [], _sr = [], _dl = true;
            if (_s && _s.length > 0) {
                $.each(_s, function (_y, _z) {
                    _ss.push(_z.DestinationPointId);
                    _sr.push(_z.RangeId);
                    _dl = _z.Deleted;
                });
                that.utility.mode(function () {
                    that.popup.confirm({
                        title: "Do you want to delete the selected destination point?",
                        message: "Removing these destination points cannot be undone."
                    }, function () {
                        // Ask if user want to stay value in document
                        that.action.requestValueInDoc(_dl, function (keepContent) {

                            that.popup.processing(true);
                            that.service.deleteSelected({ data: { "": _ss } }, function (result) {
                                if (result.status == app.status.succeeded) {
                                    that.range.delSelected({ data: _sr, index: 0 }, function () {
                                        that.popup.message({ success: true, title: "Delete destination point succeeded." }, function () { that.popup.hide(3000); });
                                        $.each(_ss, function (_m, _n) {
                                            that.utility.remove({ Id: _n });
                                            that.ui.remove({ Id: _n });
                                        });
                                        that.utility.unSelectAll();
                                        that.utility.pager.init({ refresh: false, index: that.controls.list.find(".point-item").length > 0 ? that.pagerIndex : that.pagerIndex - 1 });
                                    }, keepContent); // Add the parameter here to stay value or not.
                                }
                                else {
                                    that.popup.message({ success: false, title: "Delete destination point failed." });
                                }
                            });

                        });

                    }, function () {
                        that.controls.popupMain.removeClass("message process confirm active");
                    });
                });
            }
            else {
                that.popup.message({ success: false, title: "Please select destination point." });
            }
        },
        // Ask if user want to stay value in document
        requestValueInDoc: function (deleted, callback) {
            if (deleted) {
                callback(true);
            }
            else {
                that.popup.confirm({ title: "Do you want the value to stay in the document?", message: "" },
                    function () {
                        callback(true); // Stay value in document.
                    },
                    function () {
                        callback(false); // Not Stay value in document. 
                    });
            }
        },
        edit: function (o) {
            that.utility.mode(function () {
                var _i = $(o).data("id");
                that.model = that.utility.model(_i);
                if (that.model) {
                    that.range.goto({ RangeId: $(o).data("range") }, function (result) {
                        if (result.status == app.status.succeeded) {
                            //that.controls.main.removeClass("manage add edit clone step-first step-second step-third step-fourth").addClass("add edit");
                            if ($(o).find(".add-point-customformat").hasClass("show")) {
                                $(o).find(".add-point-customformat").removeClass("show");
                            }
                            else {
                                $(o).find(".add-point-customformat").addClass("show");
                                that.ui.status({ next: false, cancel: true, save: true });
                                $(o).find(".btnSelectFormat").prop("original", that.model.ReferencedSourcePoint.Value ? that.model.ReferencedSourcePoint.Value : "");
                                that.action.customFormat({ selected: that.model, o: o, ref: true }, function () {
                                    that.format.preview(o);
                                });
                            }
                        }
                        else {
                            that.popup.message({ success: false, title: "The point in Word has been deleted." });
                        }
                    });
                }
                else {
                    that.popup.message({ success: false, title: "The destination point has been deleted." });
                }
            });
        },
        history: function (o) {
            o.hasClass("item-more") ? o.removeClass("item-more") : o.addClass("item-more");
        },
        increase: function (o) {
            var _p = o.closest("li").find(".btnSelectFormat").prop("place"), _v = that.format.remove(o.closest("li").find(".lbPreviewValue").text());
            if (_p == "") {
                _p = that.format.getDecimalLength(_v);
            }
            _p = parseInt(_p);
            o.closest("li").find(".btnSelectFormat").prop("place", ++_p);
            that.format.preview(o.closest("li"));
        },
        decrease: function (o) {
            var _p = o.closest("li").find(".btnSelectFormat").prop("place"), _v = that.format.remove(o.closest("li").find(".lbPreviewValue").text());
            if (_p == "") {
                _p = that.format.getDecimalLength(_v);
            }
            _p = parseInt(_p);
            if (_p > 0) {
                o.closest("li").find(".btnSelectFormat").prop("place", --_p);
                that.format.preview(o.closest("li"));
            }
        },
        goto: function (o) {
            that.utility.mode(function () {
                var _selectedItem = o.hasClass("point-item") ? o : o.closest(".point-item");
                that.controls.list.find(".point-item.selected").removeClass("selected");
                _selectedItem.addClass("selected");

                var _rid = _selectedItem.data("range");
                that.range.goto({ RangeId: _rid }, function (result) {
                    if (result.status == app.status.failed) {
                        that.popup.message({ success: false, title: "The point in Word has been deleted." });
                    }
                });
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
                    that.action.goto(that.controls.list.find(">li").eq(_i));
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
                    that.action.goto(that.controls.list.find(">li").eq(_i));
                }
            }
        },
        ok: function () {
            that.controls.popupMain.removeClass("active message process confirm");
        },
        open: function (o) {
            var _p = $(o).data("path");
            if (_p) {
                window.open(_p);
            }
        },
        autoComplete: function () {
            if (that.file) {
                var _e = that.utility.entered(), _d = that.file.SourcePoints;
                if ($.trim(_e.keyword) != "") {
                    app.search.autoComplete({ keyword: _e.keyword, data: _d, result: that.controls.autoCompleteControl, target: that.controls.keyword });
                }
                else {
                    that.controls.autoCompleteControl.hide();
                }
            }
        },
        autoComplete2: function () {
            var _e = $.trim(that.controls.sourcePointName.val()), _d = that.points, _da = [];
            if (_e != "") {
                $.each(_d, function (i, d) {
                    _da.push(d.ReferencedSourcePoint);
                });
                app.search.autoComplete({ keyword: _e, data: _da, result: that.controls.autoCompleteControl2, target: that.controls.sourcePointName });
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
        selectedFormats: function (_that) {
            var _fi = [], _fd = [], _fn = [];
            _that.closest(".listFormats").find("ul > li").each(function (i, d) {
                if ($(this).hasClass("checked")) {
                    _fi.push($(this).data("id"));
                    _fd.push($.trim($(this).data("displayname")));
                    _fn.push($.trim($(this).data("name")));
                    if ($.trim($(_that).data("name")).indexOf("ConvertTo") > -1) {
                        _that.closest(".point-item").find(".btnSelectFormat").prop("place", "");
                    }
                }
            });
            _that.closest(".point-item").find(".btnSelectFormat").html(_fd.length > 0 ? _fd.join(", ") : "None");
            _that.closest(".point-item").find(".btnSelectFormat").prop("title", _fd.length > 0 ? _fd.join(", ") : "None");
            _that.closest(".point-item").find(".btnSelectFormat").prop("selected", _fi.join(","));
            _that.closest(".point-item").find(".btnSelectFormat").prop("name", _fn.join(","));
            that.format.preview($(_that).closest(".point-item"));
        },
        customFormat: function (options, callback) {
            var _r = false;
            var _pp = false;
            setTimeout(function () {
                if (!_r) {
                    that.popup.processing(true);
                    _pp = true;
                }
            }, 250);

            that.service.customFormat(function (result) {
                _r = true;
                if (_pp) {
                    that.popup.processing(false);
                }
                if (result.status == app.status.succeeded) {
                    if (result.data) {
                        that.ui.customFormat({ o: options.o, data: result.data, selected: options.selected ? options.selected : null, selectedPoint: options.selectedPoint ? options.selectedPoint : null, ref: options.ref }, callback);
                    }
                }
                else {
                    that.ui.customFormat({ o: options.o, selected: options.selected ? options.selected : null, selectedPoint: options.selectedPoint ? options.selectedPoint : null, ref: options.ref }, callback);
                    that.popup.message({ success: false, title: "Load custom format failed." });
                }
            });
        },
        selectRecentFile: function (options) {
            $(that.controls.recentFilesDrop[0]).text(options.elem.text());
            that.controls.recentFiles.find("li.is-selected").removeClass("is-selected");
            options.elem.addClass("is-selected");
            var _i = options.elem.index();
            that.file = that.recentFiles[_i];
            that.ui.select();
            that.selected = null;
            that.controls.sourceTypeNav.removeClass("is-selected");
            that.ui.sources({ data: that.file, keyword: "", sourceType: that.utility.selectedSourceType(".point-types") });
            that.controls.stepFirstMain.addClass("selected-file");
            that.utility.height();
        },
        clone: {
            source: null,
            destination: null,
            checkResult: null,
            init: function () {
                that.ui.clear();
                that.ui.dft();
                that.action.clone.source = null;
                that.action.clone.destination = null;
                that.action.clone.checkResult = null;
                that.controls.cloneNext.removeClass("ms-Button--primary").addClass("disabled");
                that.controls.cloneResult.removeClass("clone-success clone-error");
                that.controls.main.removeClass("manage add edit clone step-first step-second step-third step-fourth").addClass("add clone step-first");
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
            var _sd = elem.hasClass("sort-desc");
            that.controls.headerListPoints.find(".i2.sort-desc,.i2.sort-asc,.i3.sort-desc,.i3.sort-asc,.i4.sort-desc,.i4.sort-asc").removeClass("sort-asc sort-desc");
            elem.addClass(_sd ? "sort-asc" : "sort-desc");
            that.ui.list({ refresh: false });
        },
        resort: function () {
            that.controls.headerListPoints.find(".i2.sort-desc,.i2.sort-asc,.i3.sort-desc,.i3.sort-asc,.i4.sort-desc,.i4.sort-asc").removeClass("sort-asc sort-desc");
            that.ui.list({ refresh: false });
        },
        sortAdd: function (elem) {
            var _sd = elem.hasClass("sort-desc");
            that.controls.headerListPointsAdd.find(".i2.sort-desc,.i2.sort-asc,.i3.sort-desc,.i3.sort-asc,.i4.sort-desc,.i4.sort-asc").removeClass("sort-asc sort-desc");
            elem.addClass(_sd ? "sort-asc" : "sort-desc");
            that.ui.sources({ data: that.file, keyword: that.keyword, sourceType: that.utility.selectedSourceType(".point-types") });
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
                    var _fName = "";
                    $.each(result.data.value, function (i, d) {
                        if (decodeURI(options.url).toUpperCase() == decodeURI(d.EncodedAbsUrl).toUpperCase() && d.OData__dlc_DocId) {
                            _d = d.OData__dlc_DocId;
                            _fName = d.FileLeafRef;
                            return false;
                        }
                    });
                    if (_d != "") {
                        that.service.addFile({ data: { DocumentId: _d, Name: _fName } }, function (resultIn) {
                            if (resultIn.status == app.status.succeeded) {
                                that.popup.processing(false);
                                that.recentFiles = resultIn.data;
                                that.ui.recentFiles({ data: resultIn.data });

                                that.browse.popup.hide();
                                that.files($.extend({}, { documentId: _d }, options));
                            }
                            else {
                                that.popup.message({ success: false, title: "Load recent files failed." });
                            }

                        });
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

    that.format = {
        convert: function (options) {
            var _t = $.trim(options.value != null ? options.value : ""),
                _v = _t,
                _f = options.formats ? options.formats : [],
                _d = that.format.hasDollar(_v),
                _c = true, //that.format.hasComma(_v),
                _p = that.format.hasPercent(_v),
                _k = that.format.hasParenthesis(_v),
                _m = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
                _x = options.decimal;
            $.each(_f, function (_a, _b) {
                if (!_b.IsDeleted) {
                    if (_b.Name == "ConvertToHundreds") {
                        if (that.format.isNumber(_v)) {
                            var _l = that.format.getDecimalLength(_v);
                            _v = new BigNumber(that.format.toNumber(_v)).div(100).toString();
                            _v = that.format.addDecimal(_v, _l);
                            if (_c) {
                                _v = that.format.addComma(_v);
                            }
                            if (_p) {
                                _v = that.format.addPercent(_v);
                            }
                            if (_k) {
                                _v = "(" + _v + ")";
                            }
                            if (_d) {
                                _v = that.format.addDollar(_v);
                            }
                        }
                    }
                    else if (_b.Name == "ConvertToThousands") {
                        if (that.format.isNumber(_v)) {
                            var _l = that.format.getDecimalLength(_v);
                            _v = new BigNumber(that.format.toNumber(_v)).div(1000).toString();
                            _v = that.format.addDecimal(_v, _l);
                            if (_c) {
                                _v = that.format.addComma(_v);
                            }
                            if (_p) {
                                _v = that.format.addPercent(_v);
                            }
                            if (_k) {
                                _v = "(" + _v + ")";
                            }
                            if (_d) {
                                _v = that.format.addDollar(_v);
                            }
                        }
                    }
                    else if (_b.Name == "ConvertToMillions") {
                        if (that.format.isNumber(_v)) {
                            var _l = that.format.getDecimalLength(_v);
                            _v = new BigNumber(that.format.toNumber(_v)).div(1000000).toString();
                            _v = that.format.addDecimal(_v, _l);
                            if (_c) {
                                _v = that.format.addComma(_v);
                            }
                            if (_p) {
                                _v = that.format.addPercent(_v);
                            }
                            if (_k) {
                                _v = "(" + _v + ")";
                            }
                            if (_d) {
                                _v = that.format.addDollar(_v);
                            }
                        }
                    }
                    else if (_b.Name == "ConvertToBillions") {
                        if (that.format.isNumber(_v)) {
                            var _l = that.format.getDecimalLength(_v);
                            _v = new BigNumber(that.format.toNumber(_v)).div(1000000000).toString();
                            _v = that.format.addDecimal(_v, _l);
                            if (_c) {
                                _v = that.format.addComma(_v);
                            }
                            if (_p) {
                                _v = that.format.addPercent(_v);
                            }
                            if (_k) {
                                _v = "(" + _v + ")";
                            }
                            if (_d) {
                                _v = that.format.addDollar(_v);
                            }
                        }
                    }
                    else if (_b.Name == "ShowNegativesAsPositives") {
                        var _h = that.format.hasDollar(_v),
                            _pt = that.format.hasPercent(_v),
                            _pk = that.format.hasParenthesis(_v),
                            _hh = _v.toString().indexOf("hundred") > -1,
                            _ht = _v.toString().indexOf("thousand") > -1,
                            _hm = _v.toString().indexOf("million") > -1,
                            _hb = _v.toString().indexOf("billion") > -1;
                        _v = $.trim(_v.toString().replace(/\$/g, "").replace(/-/g, "").replace(/%/g, "").replace(/\(/g, "").replace(/\)/g, "").replace(/hundred/g, "").replace(/thousand/g, "").replace(/million/g, "").replace(/billion/g, ""));
                        if (_pt) {
                            _v = that.format.addPercent(_v);
                        }
                        if (_h) {
                            _v = that.format.addDollar(_v);
                        }
                        if (_hh) {
                            _v = _v + " hundred";
                        }
                        else if (_ht) {
                            _v = _v + " thousand";
                        }
                        else if (_hm) {
                            _v = _v + " million";
                        }
                        else if (_hb) {
                            _v = _v + " billion";
                        }
                    }
                    else if (_b.Name == "IncludeHundredDescriptor") {
                        if (that.format.isNumber(_v)) {
                            _v = _v + " hundred";
                        }
                    }
                    else if (_b.Name == "IncludeThousandDescriptor") {
                        if (that.format.isNumber(_v)) {
                            _v = _v + " thousand";
                        }
                    }
                    else if (_b.Name == "IncludeMillionDescriptor") {
                        if (that.format.isNumber(_v)) {
                            _v = _v + " million";
                        }
                    }
                    else if (_b.Name == "IncludeBillionDescriptor") {
                        if (that.format.isNumber(_v)) {
                            _v = _v + " billion";
                        }
                    }
                    else if (_b.Name == "IncludeDollarSymbol") {
                        if (!that.format.hasDollar(_v)) {
                            _v = that.format.addDollar(_v);
                        }
                    }
                    else if (_b.Name == "ExcludeDollarSymbol") {
                        if (that.format.hasDollar(_v)) {
                            _v = that.format.removeDollar(_v);
                        }
                    }
                    else if (_b.Name == "DateShowLongDateFormat") {
                        if (that.format.isDate(_v)) {
                            var _tt = new Date(_v);
                            _v = _m[_tt.getMonth()] + " " + _tt.getDate() + ", " + _tt.getFullYear();
                        }
                    }
                    else if (_b.Name == "DateShowYearOnly") {
                        if (that.format.isDate(_v)) {
                            var _tt = new Date(_v);
                            _v = _tt.getFullYear();
                        }
                    }
                    else if (_b.Name == "ConvertNegativeSymbolToParenthesis") {
                        var _h = that.format.hasDollar(_v),
                            _pt = that.format.hasPercent(_v),
                            _hh = _v.toString().indexOf("hundred") > -1,
                            _ht = _v.toString().indexOf("thousand") > -1,
                            _hm = _v.toString().indexOf("million") > -1,
                            _hb = _v.toString().indexOf("billion") > -1;
                        if (_v.indexOf("-") > -1) {
                            _v = $.trim(_v.toString().replace(/\$/g, "").replace(/-/g, "").replace(/%/g, "").replace(/\(/g, "").replace(/\)/g, "").replace(/hundred/g, "").replace(/thousand/g, "").replace(/million/g, "").replace(/billion/g, ""));
                            if (_h) {
                                _v = that.format.addDollar(_v);
                            }
                            _v = "(" + _v + ")";
                            if (_pt) {
                                _v = that.format.addPercent(_v);
                            }
                            if (_hh) {
                                _v = _v + " hundred";
                            }
                            else if (_ht) {
                                _v = _v + " thousand";
                            }
                            else if (_hm) {
                                _v = _v + " million";
                            }
                            else if (_hb) {
                                _v = _v + " billion";
                            }
                        }
                    }
                }
            });
            if (_x != null && _x.toString() != "") {
                var __a = that.format.hasComma(_v), __b = that.format.remove(_v);
                if (that.format.isNumber(__b)) {
                    var __c = __a ? that.format.addComma(__b) : __b, __d = "" + new BigNumber(__b).toFixed(_x) + "", __e = __a ? that.format.addComma(__d) : __d;
                    _v = _v.replace(__c, __e)
                }
            }
            return _v;
        },
        toNumber: function (_v) {
            return that.format.removeComma(that.format.removeDollar(that.format.removePercent(that.format.removeParenthesis(_v))));
        },
        isNumber: function (_v) {
            return that.format.toNumber(_v) != "" && !isNaN(that.format.toNumber(_v));
        },
        isDate: function (_v) {
            var _ff = false;
            try {
                _ff = _ff = (new Date(_v.toString().replace(/ /g, ""))).getFullYear() > 0;
            } catch (e) {
                _ff = false;
            }
            return _ff;
        },
        hasDollar: function (_v) {
            return _v.toString().indexOf("$") > -1;
        },
        hasComma: function (_v) {
            return _v.toString().indexOf(",") > -1;
        },
        removeDollar: function (_v) {
            return _v.toString().replace("$", "");
        },
        addDollar: function (_v) {
            return "$" + _v;
        },
        removeComma: function (_v) {
            return _v.toString().replace(/,/g, "");
        },
        addComma: function (_v) {
            var __s = _v.toString().split(".");
            __s[0] = __s[0].replace(new RegExp('(\\d)(?=(\\d{3})+$)', 'ig'), "$1,");
            return __s.join(".");
        },
        removePercent: function (_v) {
            return _v.toString().replace(/%/g, "");
        },
        hasPercent: function (_v) {
            return _v.toString().indexOf("%") > -1;
        },
        addPercent: function (_v) {
            return _v + "%";
        },
        hasParenthesis: function (_v) {
            return _v.toString().indexOf("(") > -1 && _v.toString().indexOf(")") > -1;
        },
        removeParenthesis: function (_v) {
            return _v.toString().replace(/\(/g, "").replace(/\)/g, "");
        },
        remove: function (_v) {
            return $.trim(_v.toString().replace(/\$/g, "").replace(/,/g, "").replace(/-/g, "").replace(/%/g, "").replace(/\(/g, "").replace(/\)/g, "").replace(/hundred/g, "").replace(/thousand/g, "").replace(/million/g, "").replace(/billion/g, ""));
        },
        getDecimalLength: function (_v) {
            var _a = _v.toString().replace(/\$/g, "").replace(/,/g, "").replace(/-/g, "").replace(/\(/g, "").replace(/\)/g, "").split(".");
            if (_a.length == 2) {
                return _a[1].length;
            }
            else {
                return 0;
            }
        },
        addDecimal: function (_v, _l) {
            var _dl = that.format.getDecimalLength(_v);
            if (_l > 0 && _dl == 0) {
                _v = "" + new BigNumber(_v).toFixed(_l) + "";
            }
            return _v;
        },
        preview: function (o) {
            var _v = o.find(".btnSelectFormat").prop("original");
            var _n = o.find(".btnSelectFormat").prop("name");
            var _f = (typeof (_n) != "undefined" && _n != "") ? _n.split(",") : [], _fa = [];
            var _x = o.find(".btnSelectFormat").prop("place");
            $.each(_f, function (_a, _b) {
                _fa.push({ Name: _b });
            });
            var _fd = that.format.convert({ value: _v, formats: _fa, decimal: _x });
            o.find(".lbPreviewValue").html(_fd);
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
            that.controls.popupConfirmMessage.html(options.message);
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
                    that.controls.main.removeClass("manage add edit clone step-first step-third step-fourth").addClass("manage");
                }, millisecond);
            }
            else {
                $(".popups .bg").removeAttr("style");
                that.controls.popupMain.removeClass("active message");
                that.controls.main.removeClass("manage add edit clone step-first step-third step-fourth").addClass("manage");
            }
        }
    };

    that.range = {
        create: function (options, callback) {
            Word.run(function (ctx) {
                var _r = ctx.document.getSelection(), _cc = _r.insertContentControl();
                _cc.tag = options.RangeId;
                _cc.title = options.Name;
                ctx.load(_cc);
                return ctx.sync().then(function () {
                    if (options.DestinationType == app.destinationTypes.point) {
                        _cc.insertText(app.string(options.Value), Word.InsertLocation.replace);
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded });
                        });
                    }
                    else if (options.DestinationType == app.destinationTypes.tableCell) {
                        var _d = JSON.parse(options.Value);
                        var _v = _d.table.values, _f = _d.table.formats, _i = 0;
                        var _table = _cc.insertTable(_v.length, _v[0].length, Word.InsertLocation.start, _v);
                        _table.load("rows/cells");
                        return ctx.sync().then(function () {
                            for (var i = 0; i < _table.rows.items.length; i++) {
                                if (_f[_i].preferredHeight) {
                                    var _row = _table.rows.items[i];
                                    _row.preferredHeight = _f[_i].preferredHeight;
                                }
                                var _cells = _table.rows.items[i].cells.items;
                                for (var j = 0; j < _cells.length; j++) {
                                    _cells[j].columnWidth = _f[_i].columnWidth;
                                    _cells[j].horizontalAlignment = _f[_i].horizontalAlignment;
                                    _cells[j].verticalAlignment = _f[_i].verticalAlignment;
                                    _cells[j].shadingColor = _f[_i].shadingColor;
                                    _cells[j].body.font.bold = _f[_i].font.bold;
                                    _cells[j].body.font.color = _f[_i].font.color;
                                    _cells[j].body.font.italic = _f[_i].font.italic;
                                    _cells[j].body.font.name = _f[_i].font.name;
                                    _cells[j].body.font.size = _f[_i].font.size;
                                    _cells[j].body.font.underline = _f[_i].font.underline;
                                    _cells[j].getBorder(Word.BorderLocation.top).color = _f[_i].border.top.color;
                                    _cells[j].getBorder(Word.BorderLocation.top).type = _f[_i].border.top.type;
                                    _cells[j].getBorder(Word.BorderLocation.bottom).color = _f[_i].border.bottom.color;
                                    _cells[j].getBorder(Word.BorderLocation.bottom).type = _f[_i].border.bottom.type;
                                    _cells[j].getBorder(Word.BorderLocation.left).color = _f[_i].border.left.color;
                                    _cells[j].getBorder(Word.BorderLocation.left).type = _f[_i].border.left.type;
                                    _cells[j].getBorder(Word.BorderLocation.right).color = _f[_i].border.right.color;
                                    _cells[j].getBorder(Word.BorderLocation.right).type = _f[_i].border.right.type;
                                    //cells[j].getBorder(Word.BorderLocation.insideVertical).color = _f[_i].border.insideVertical.color;
                                    //cells[j].getBorder(Word.BorderLocation.insideVertical).type = _f[_i].border.insideVertical.type;
                                    //cells[j].getBorder(Word.BorderLocation.insideHorizontal).color = _f[_i].border.insideHorizontal.color;
                                    //cells[j].getBorder(Word.BorderLocation.insideHorizontal).type = _f[_i].border.insideHorizontal.type;
                                    _i++;
                                }
                            }
                            return ctx.sync().then(function () {
                                var _pg = _cc.paragraphs;
                                ctx.load(_pg);
                                return ctx.sync().then(function () {
                                    if (_pg.items.length > 0) {
                                        _pg.items[0].delete();
                                        return ctx.sync().then(function () {
                                            callback({ status: app.status.succeeded });
                                        });
                                    }
                                    else {
                                        callback({ status: app.status.succeeded });
                                    }
                                });
                            });
                        });
                    }
                    else if (options.DestinationType == app.destinationTypes.tableImage) {
                        var _d = JSON.parse(options.Value);
                        _cc.insertInlinePictureFromBase64(_d.image, Word.InsertLocation.replace);
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded });
                        });
                    }
                    else {
                        $("#debugInfo").append("Name:" + options.Name + ", Value:" + options.Value + "<br/>");
                        _cc.insertInlinePictureFromBase64(options.Value, Word.InsertLocation.replace);
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded });
                        });
                    }
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        goto: function (options, callback) {
            Word.run(function (ctx) {
                var _cc = ctx.document.contentControls.getByTag(options.RangeId).getFirst();
                ctx.load(_cc, "items");
                return ctx.sync().then(function () {
                    _cc.select();
                    callback({ status: app.status.succeeded });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        exist: function (options, callback) {
            Word.run(function (ctx) {
                var _cc = ctx.document.contentControls.getByTag(options.RangeId).getFirst();
                ctx.load(_cc);
                return ctx.sync().then(function () {
                    callback({ status: app.status.succeeded });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        highlight: function (options, callback) {
            Word.run(function (ctx) {
                var _cc = ctx.document.contentControls.getByTag(options.RangeId).getFirst();
                ctx.load(_cc);
                return ctx.sync().then(function () {
                    _cc.font.highlightColor = that.highlighted ? "" : that.highlightColor;
                    return ctx.sync().then(function () {
                        callback({ status: app.status.succeeded });
                    });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        edit: function (options, callback) {
            Word.run(function (ctx) {
                var _cc = ctx.document.contentControls.getByTag(options.RangeId).getFirst();
                ctx.load(_cc, "text");
                return ctx.sync().then(function () {
                    var _t = app.string(_cc.text), _v = app.string(options.Value);
                    if (_t != _v) {
                        _cc.insertText(_v, Word.InsertLocation.replace);
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded });
                        });
                    }
                    else {
                        callback({ status: app.status.succeeded });
                    }
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        editTable: function (options, callback) {
            Word.run(function (ctx) {
                var _cc = ctx.document.contentControls.getByTag(options.RangeId).getFirst();
                ctx.load(_cc);
                return ctx.sync().then(function () {
                    var _table = _cc.tables.getFirst();
                    ctx.load(_table);
                    return ctx.sync().then(function () {
                        _table.delete();
                        return ctx.sync().then(function () {
                            var _d = JSON.parse(options.Value), _v = _d.values, _f = _d.formats, _i = 0;
                            var _newTable = _cc.insertTable(_v.length, _v[0].length, Word.InsertLocation.start, _v);
                            _newTable.load("rows/cells");
                            return ctx.sync().then(function () {
                                for (var i = 0; i < _newTable.rows.items.length; i++) {
                                    if (_f[_i].preferredHeight) {
                                        var _row = _newTable.rows.items[i];
                                        _row.preferredHeight = _f[_i].preferredHeight;
                                    }
                                    var _cells = _newTable.rows.items[i].cells.items;
                                    for (var j = 0; j < _cells.length; j++) {
                                        _cells[j].columnWidth = _f[_i].columnWidth;
                                        _cells[j].horizontalAlignment = _f[_i].horizontalAlignment;
                                        _cells[j].verticalAlignment = _f[_i].verticalAlignment;
                                        _cells[j].shadingColor = _f[_i].shadingColor;
                                        _cells[j].body.font.bold = _f[_i].font.bold;
                                        _cells[j].body.font.color = _f[_i].font.color;
                                        _cells[j].body.font.italic = _f[_i].font.italic;
                                        _cells[j].body.font.name = _f[_i].font.name;
                                        _cells[j].body.font.size = _f[_i].font.size;
                                        _cells[j].body.font.underline = _f[_i].font.underline;
                                        _cells[j].getBorder(Word.BorderLocation.top).color = _f[_i].border.top.color;
                                        _cells[j].getBorder(Word.BorderLocation.top).type = _f[_i].border.top.type;
                                        _cells[j].getBorder(Word.BorderLocation.bottom).color = _f[_i].border.bottom.color;
                                        _cells[j].getBorder(Word.BorderLocation.bottom).type = _f[_i].border.bottom.type;
                                        _cells[j].getBorder(Word.BorderLocation.left).color = _f[_i].border.left.color;
                                        _cells[j].getBorder(Word.BorderLocation.left).type = _f[_i].border.left.type;
                                        _cells[j].getBorder(Word.BorderLocation.right).color = _f[_i].border.right.color;
                                        _cells[j].getBorder(Word.BorderLocation.right).type = _f[_i].border.right.type;
                                        _i++;
                                    }
                                }
                                return ctx.sync().then(function () {
                                    var _pg = _cc.paragraphs;
                                    ctx.load(_pg);
                                    return ctx.sync().then(function () {
                                        if (_pg.items.length > 0) {
                                            _pg.items[0].delete();
                                            return ctx.sync().then(function () {
                                                callback({ status: app.status.succeeded });
                                            });
                                        }
                                        else {
                                            callback({ status: app.status.succeeded });
                                        }
                                    });
                                });
                            });
                        });
                    });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        del: function (options, callback, keepContent) {
            Word.run(function (ctx) {
                var _cc = ctx.document.contentControls.getByTag(options.RangeId);
                ctx.load(_cc);
                return ctx.sync().then(function () {
                    if (_cc.items.length > 0) {
                        _cc.items[0].delete(keepContent);
                        return ctx.sync().then(function () {
                            callback({ status: app.status.succeeded });
                        });
                    }
                    else {
                        callback({ status: app.status.succeeded });
                    }
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        delSelected: function (options, callback, keepContent) {
            if (options.index < options.data.length) {
                var _rid = options.data[options.index];
                Word.run(function (ctx) {
                    var _cc = ctx.document.contentControls.getByTag(_rid);
                    ctx.load(_cc);
                    return ctx.sync().then(function () {
                        if (_cc.items.length > 0) {
                            _cc.items[0].delete(keepContent);
                            return ctx.sync().then(function () {
                                options.index++;
                                that.range.delSelected(options, callback, keepContent);
                            });
                        }
                        else {
                            options.index++;
                            that.range.delSelected(options, callback, keepContent);
                        }
                    });
                }).catch(function (error) {
                    options.index++;
                    that.range.delSelected(options, callback, keepContent);
                });
            }
            else {
                callback();
            }
        },
        all: function (callback) {
            Word.run(function (ctx) {
                var _cc = ctx.document.contentControls, _ar = [];
                ctx.load(_cc, "tag");
                return ctx.sync().then(function () {
                    if (_cc.items.length > 0) {
                        for (var _i = 0; _i < _cc.items.length; _i++) {
                            _ar.push({ tag: _cc.items[_i].tag });
                        }
                    }
                    callback({ status: app.status.succeeded, data: _ar });
                });
            }).catch(function (error) {
                callback({ status: app.status.failed, message: error.message });
            });
        },
        index: function (options) {
            var _i = -1;
            $.each(options.data, function (i, d) {
                if (d.tag == options.tag) {
                    _i = i;
                    return false;
                }
            });
            return _i;
        },
        valueItem: function (controlItems, options, ctx, callback) {
            var handleError = function () {
                options.data[options.index].existed = false;
                options.data[options.index].changed = false;
                options.data[options.index].orderBy = 99999;
                options.data[options.index].PublishedStatus = false;
                options.index++;
                that.range.valueItem(controlItems, options, ctx, callback);
            };
            if (options.index >= options.data.length) {
                callback({ data: options.data });
                return;
            }
            var _dsp = options.data[options.index],
           _sp = _dsp.ReferencedSourcePoint,
           _cf = _dsp.CustomFormats,
           _dp = _dsp.DecimalPlace,
           _tag = _dsp.RangeId,
           _pvt = _sp.PublishedHistories && _sp.PublishedHistories.length > 0 ? (_sp.PublishedHistories[0].Value ? _sp.PublishedHistories[0].Value : "") : "",
           _pv = _sp.SourceType == app.sourceTypes.table ? (_dsp.DestinationType == app.destinationTypes.tableImage ? JSON.parse(_pvt).image : JSON.stringify(JSON.parse(_pvt).table)) : _pvt,
           _index = that.range.index({ data: options.tags, tag: _tag });
            var _cc;
            for (var i = 0; i < controlItems.length; i++) {
                if (_tag === controlItems[i].tag) {
                    _cc = controlItems[i];
                    break;
                }
            }
            if (!_cc) {
                handleError();
            }
            else if (_sp.SourceType == app.sourceTypes.point) {
                var _v = that.format.convert({ value: app.string(_sp.Value), formats: _cf, decimal: _dp }),
                    _c = _cc.text != _v;
                options.data[options.index].existed = true;
                options.data[options.index].changed = _c;
                options.data[options.index].orderBy = _index > -1 ? _index : 99999;
                options.data[options.index].PublishedStatus = options.refresh ? true : (_v == _pv);
                if (options.refresh && _c) {
                    _cc.insertText(_v, Word.InsertLocation.replace);
                    return ctx.sync().then(function () {
                        options.index++;
                        that.range.valueItem(controlItems, options, ctx, callback);
                    });
                }
                else {
                    options.index++;
                    that.range.valueItem(controlItems, options, ctx, callback);
                }
            }
            else if (_sp.SourceType == app.sourceTypes.table && _dsp.DestinationType != app.destinationTypes.tableImage) {
                var _table = _cc.tables.getFirst();
                _table.load("values");
                return ctx.sync().then(function () {
                    that.range.formats({ tag: _dsp.RangeId, data: _table.values }, function (result) {
                        if (result.status == app.status.succeeded) {
                            var _v = JSON.parse(_sp.Value).table,
                                _c = result.data != JSON.stringify(_v);
                            options.data[options.index].existed = true;
                            options.data[options.index].changed = _c;
                            options.data[options.index].orderBy = _index > -1 ? _index : 99999;
                            options.data[options.index].PublishedStatus = options.refresh ? true : (result.data == _pv);
                            if (options.refresh && _c) {
                                _table.delete();
                                return ctx.sync().then(function () {
                                    var _f = _v.formats, _tv = _v.values, _i = 0;
                                    var _newTable = _cc.insertTable(_tv.length, _tv[0].length, Word.InsertLocation.start, _tv);
                                    _newTable.load("rows/cells");
                                    return ctx.sync().then(function () {
                                        for (var i = 0; i < _newTable.rows.items.length; i++) {
                                            var _cells = _newTable.rows.items[i].cells.items;
                                            for (var j = 0; j < _cells.length; j++) {
                                                _cells[j].columnWidth = _f[_i].columnWidth;
                                                _cells[j].horizontalAlignment = _f[_i].horizontalAlignment;
                                                _cells[j].verticalAlignment = _f[_i].verticalAlignment;
                                                _cells[j].shadingColor = _f[_i].shadingColor;
                                                _cells[j].body.font.bold = _f[_i].font.bold;
                                                _cells[j].body.font.color = _f[_i].font.color;
                                                _cells[j].body.font.italic = _f[_i].font.italic;
                                                _cells[j].body.font.name = _f[_i].font.name;
                                                _cells[j].body.font.size = _f[_i].font.size;
                                                _cells[j].body.font.underline = _f[_i].font.underline;
                                                _cells[j].getBorder(Word.BorderLocation.top).color = _f[_i].border.top.color;
                                                _cells[j].getBorder(Word.BorderLocation.top).type = _f[_i].border.top.type;
                                                _cells[j].getBorder(Word.BorderLocation.bottom).color = _f[_i].border.bottom.color;
                                                _cells[j].getBorder(Word.BorderLocation.bottom).type = _f[_i].border.bottom.type;
                                                _cells[j].getBorder(Word.BorderLocation.left).color = _f[_i].border.left.color;
                                                _cells[j].getBorder(Word.BorderLocation.left).type = _f[_i].border.left.type;
                                                _cells[j].getBorder(Word.BorderLocation.right).color = _f[_i].border.right.color;
                                                _cells[j].getBorder(Word.BorderLocation.right).type = _f[_i].border.right.type;
                                                _i++;
                                            }
                                        }
                                        return ctx.sync().then(function () {
                                            var _pg = _cc.paragraphs;
                                            ctx.load(_pg);
                                            return ctx.sync().then(function () {
                                                if (_pg.items.length > 0) {
                                                    _pg.items[0].delete();
                                                    return ctx.sync().then(function () {
                                                        options.index++;
                                                        that.range.valueItem(controlItems, options, ctx, callback);
                                                    });
                                                }
                                                else {
                                                    options.index++;
                                                    that.range.valueItem(controlItems, options, ctx, callback);
                                                }
                                            }).catch(function (error) {
                                                handleError();
                                            });
                                        }).catch(function (error) {
                                            handleError();
                                        });
                                    }).catch(function (error) {
                                        handleError();
                                    });
                                }).catch(function (error) {
                                    handleError();
                                });
                            }
                            else {
                                options.index++;
                                that.range.valueItem(controlItems, options, ctx, callback);
                            }
                        }
                        else {
                            options.data[options.index].existed = false;
                            options.data[options.index].changed = false;
                            options.data[options.index].orderBy = 99999;
                            options.data[options.index].PublishedStatus = false;
                            options.index++;
                            that.range.valueItem(controlItems, options, ctx, callback);
                        }
                    });
                }).catch(function (error) {
                    handleError();
                });
            }
            else {
                var _image = _cc.inlinePictures.getFirstOrNullObject();
                ctx.load(_image, "height,width");
                return ctx.sync().then(function () {
                    var _src = _image.getBase64ImageSrc(),
                        _width = _image.width,
                        _height = _image.height;
                    return ctx.sync().then(function () {
                        var _v = _sp.SourceType == app.sourceTypes.table ? JSON.parse(_sp.Value).image : _sp.Value,
                            _c = _src.value != _v;
                        options.data[options.index].existed = true;
                        options.data[options.index].changed = _c;
                        options.data[options.index].orderBy = _index > -1 ? _index : 99999;
                        options.data[options.index].PublishedStatus = options.refresh ? true : (_v == _pv);
                        if (options.refresh && _c) {
                            _cc.insertInlinePictureFromBase64(_v, Word.InsertLocation.replace);
                            return ctx.sync().then(function () {
                                var __image = _cc.inlinePictures.getFirstOrNullObject();
                                ctx.load(__image, "height,width");
                                return ctx.sync().then(function () {
                                    __image.width = _width;
                                    return ctx.sync().then(function () {
                                        options.index++;
                                        that.range.valueItem(controlItems, options, ctx, callback);
                                    });
                                });
                            });
                        }
                        else {
                            options.index++;
                            that.range.valueItem(controlItems, options, ctx, callback);
                        }
                    });
                }).catch(function (error) {
                    handleError();
                });;
            }
        },
        values: function (options, callback) {
            if (that.firstLoad || options.refresh) {
                var handleError = function () {
                    options.data[options.index].existed = false;
                    options.data[options.index].changed = false;
                    options.data[options.index].orderBy = 99999;
                    options.data[options.index].PublishedStatus = false;
                    options.index++;
                    that.range.values(options, callback);
                };
                if (typeof (options.index) == "undefined") {
                    options.index = 0;
                }
                if (options.index < options.data.length) {
                    var _dsp = options.data[options.index],
                        _sp = _dsp.ReferencedSourcePoint,
                        _cf = _dsp.CustomFormats,
                        _dp = _dsp.DecimalPlace,
                        _tag = _dsp.RangeId,
                        _pvt = _sp.PublishedHistories && _sp.PublishedHistories.length > 0 ? (_sp.PublishedHistories[0].Value ? _sp.PublishedHistories[0].Value : "") : "",
                        _pv = _sp.SourceType == app.sourceTypes.table ? (_dsp.DestinationType == app.destinationTypes.tableImage ? JSON.parse(_pvt).image : JSON.stringify(JSON.parse(_pvt).table)) : _pvt,
                        _index = that.range.index({ data: options.tags, tag: _tag });
                    Word.run(function (ctx) {
                        var _cc = ctx.document.contentControls.getByTag(_tag).getFirst();
                        _cc.title = _sp.Name;
                        ctx.load(_cc, "text");
                        return ctx.sync().then(function () {
                            if (_sp.SourceType == app.sourceTypes.point) {
                                var _v = that.format.convert({ value: app.string(_sp.Value), formats: _cf, decimal: _dp }),
                                    _pvf = that.format.convert({ value: _pv, formats: _cf, decimal: _dp }),
                                    _c = _cc.text != _v;
                                options.data[options.index].existed = true;
                                options.data[options.index].changed = _c;
                                options.data[options.index].orderBy = _index > -1 ? _index : 99999;
                                options.data[options.index].PublishedStatus = options.refresh ? true : ($.trim(_cc.text) == _pvf);
                                options.data[options.index].DocumentValue = $.trim(_cc.text);
                                if (options.refresh && _c) {
                                    _cc.insertText(_v, Word.InsertLocation.replace);
                                    return ctx.sync().then(function () {
                                        options.index++;
                                        that.range.values(options, callback);
                                    }).catch(function (error) {
                                        handleError();
                                    });
                                }
                                else {
                                    options.index++;
                                    that.range.values(options, callback);
                                }
                            }
                            else if (_sp.SourceType == app.sourceTypes.table && _dsp.DestinationType != app.destinationTypes.tableImage) {
                                var _table = _cc.tables.getFirst();
                                _table.load("values");
                                return ctx.sync().then(function () {
                                    that.range.formats({ tag: _dsp.RangeId, data: _table.values }, function (result) {
                                        if (result.status == app.status.succeeded) {
                                            var _v = JSON.parse(_sp.Value).table,
                                                _c = result.data != JSON.stringify(_v);
                                            options.data[options.index].existed = true;
                                            options.data[options.index].changed = _c;
                                            options.data[options.index].orderBy = _index > -1 ? _index : 99999;
                                            options.data[options.index].PublishedStatus = options.refresh ? true : (result.data == _pv);
                                            if (options.refresh && _c) {
                                                that.range.editTable({ RangeId: _tag, Value: JSON.stringify(_v) }, function (result) {
                                                    if (result.status == app.status.succeeded) {
                                                        options.index++;
                                                        that.range.values(options, callback);
                                                    }
                                                    else {
                                                        handleError();
                                                    }
                                                });
                                            }
                                            else {
                                                options.index++;
                                                that.range.values(options, callback);
                                            }
                                        }
                                        else {
                                            handleError();
                                        }
                                    });
                                });
                            }
                            else {
                                var _image = _cc.inlinePictures.getFirstOrNullObject();
                                ctx.load(_image, "height,width");
                                return ctx.sync().then(function () {
                                    var _src = _image.getBase64ImageSrc(),
                                        _width = _image.width,
                                        _height = _image.height;
                                    return ctx.sync().then(function () {
                                        var _v = _sp.SourceType == app.sourceTypes.table ? JSON.parse(_sp.Value).image : _sp.Value,
                                            _c = _src.value != _v;
                                        options.data[options.index].existed = true;
                                        options.data[options.index].changed = _c;
                                        options.data[options.index].orderBy = _index > -1 ? _index : 99999;
                                        options.data[options.index].PublishedStatus = options.refresh ? true : (_src.value == _pv);
                                        if (options.refresh && _c) {
                                            $("#debugInfo").append("Name:" + _sp.Name + ", Value:" + _v + "<br/>");
                                            _cc.insertInlinePictureFromBase64(_v, Word.InsertLocation.replace);
                                            return ctx.sync().then(function () {
                                                var __image = _cc.inlinePictures.getFirstOrNullObject();
                                                ctx.load(__image, "height,width");
                                                return ctx.sync().then(function () {
                                                    __image.width = _width;
                                                    return ctx.sync().then(function () {
                                                        options.index++;
                                                        that.range.values(options, callback);
                                                    });
                                                });
                                            });
                                        }
                                        else {
                                            options.index++;
                                            that.range.values(options, callback);
                                        }
                                    });
                                });
                            }
                        });
                    }).catch(function (error) {
                        handleError();
                    });
                }
                else {
                    callback({ data: options.data });
                }
            }
            else {
                callback({ data: options.data });
            }
        },
        formats: function (options, callback) {
            if (typeof (options.index) == "undefined") {
                options.index = 0;
                options.formats = [];
                options.cells = [];
                $.each(options.data, function (m, n) {
                    $.each(n, function (x, y) {
                        options.cells.push({ row: m, column: x });
                    });
                });
            }
            if (options.index < options.cells.length) {
                Word.run(function (ctx) {
                    var _cell = ctx.document.contentControls.getByTag(options.tag).getFirst().tables.getFirst().getCell(options.cells[options.index].row, options.cells[options.index].column), _format = {};
                    var _row = _cell.parentRow;
                    _cell.load(["*", "body/font"]);
                    _row.load(["preferredHeight"]);
                    return ctx.sync().then(function () {
                        _format.columnWidth = _cell.columnWidth;
                        _format.preferredHeight = _row.preferredHeight;
                        _format.horizontalAlignment = _cell.horizontalAlignment;
                        _format.verticalAlignment = _cell.verticalAlignment;
                        _format.shadingColor = _cell.shadingColor;
                        _format.font = {
                            bold: _cell.body.font.bold,
                            color: _cell.body.font.color,
                            italic: _cell.body.font.italic,
                            name: _cell.body.font.name,
                            size: _cell.body.font.size,
                            underline: _cell.body.font.underline
                        };
                        var _top = _cell.getBorder(Word.BorderLocation.top);
                        var _left = _cell.getBorder(Word.BorderLocation.left);
                        var _right = _cell.getBorder(Word.BorderLocation.right);
                        var _bottom = _cell.getBorder(Word.BorderLocation.bottom);
                        _top.load("color,type");
                        _left.load("color,type");
                        _right.load("color,type");
                        _bottom.load("color,type");
                        return ctx.sync().then(function () {
                            _format.border = {
                                bottom: { color: _bottom.color, type: _bottom.type },
                                left: { color: _left.color, type: _left.type },
                                right: { color: _right.color, type: _right.type },
                                top: { color: _top.color, type: _top.type }
                            };
                            options.formats.push(_format);
                            options.index++;
                            that.range.formats(options, callback);
                        });
                    });
                }).catch(function (error) {
                    callback({ status: app.status.failed });
                });
            }
            else {
                callback({ status: app.status.succeeded, data: JSON.stringify({ values: options.data, formats: options.formats }) });
            }
        }
    };

    that.service = {
        common: function (options, callback) {

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
                            //console.log(that.token);
							//window.location = "/word/point";
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
        catalog: function (options, callback) {
            that.service.common({ url: that.endpoints.catalog + options.documentId, type: "GET", dataType: "json" }, callback);
        },
        list: function (callback) {
            that.service.common({ url: that.endpoints.list + "fileName=" + that.filePath + "&documentId=" + that.documentId, type: "GET", dataType: "json" }, callback);
        },
        del: function (options, callback) {
            that.service.common({ url: that.endpoints.del + options.Id, type: "DELETE" }, callback);
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
        customFormat: function (callback) {
            that.service.common({ url: that.endpoints.customFormat, type: "GET", dataType: "json" }, callback);
        },
        update: function (options, callback) {
            that.service.common({ url: that.endpoints.updateCustomFormat, type: "PUT", data: options.data, dataType: "json" }, callback);
        },
        recentFiles: function (callback) {
            that.service.common({ url: that.endpoints.recentFiles, type: "GET", dataType: "json" }, callback);
        },
        addFile: function (options, callback) {
            that.service.common({ url: that.endpoints.addRecentFile, type: "POST", data: options.data, dataType: "json" }, callback);
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
                return original.filter(function (item) {
                    var _dsp = item, _item = _dsp.ReferencedSourcePoint, _sourcePointCatalog = _item.Catalog, _s = _dsp.existed && _item.Status == 0;
                    var _fd = !_sourcePointCatalog.IsDeleted;
                    if (!_fd) {
                        return true;
                    } else {
                        if (!_s) {
                            return true;
                        }
                        else if (_item.Status === 1) {
                            return true;
                        }
                    }
                    return false;
                });
            return original;
        },
        init: function () {
            that.errorFilter.$errorFilter().click(function (event) {
                that.errorFilter.doErrorFilter($(this), event);
            });
        }
    };

    that.ui = {
        clear: function () {
            that.controls.file.val("");
            that.controls.keyword.val("");
            that.controls.sourceFolder.val("");
            that.controls.destinationFolder.val("");
        },
        dft: function () {
            var _f = $.trim(that.controls.file.val()), _fd = that.controls.file.data("default"),
                _k = $.trim(that.controls.keyword.val()), _kd = that.controls.keyword.data("default"),
                _s = $.trim(that.controls.sourceFolder.val()), _sd = that.controls.sourceFolder.data("default"),
                _d = $.trim(that.controls.destinationFolder.val()), _dd = that.controls.destinationFolder.data("default");
            if (_f == "" || _f == _fd) {
                that.controls.file.val(_fd).addClass("input-default");
            }
            if (_k == "" || _k == _kd) {
                that.controls.keyword.val(_kd).addClass("input-default");
                that.controls.search.removeClass("ms-Icon--Search ms-Icon--Cancel").addClass("ms-Icon--Search");
            }
            if (_s == "" || _s == _sd) {
                that.controls.sourceFolder.val(_sd).addClass("input-default");
            }
            if (_d == "" || _d == _dd) {
                that.controls.destinationFolder.val(_dd).addClass("input-default");
            }
        },
        status: function (options) {
            options.next ? that.controls.next.removeClass("disabled ms-Button--primary").addClass("ms-Button--primary") : that.controls.next.removeClass("ms-Button--primary").addClass("disabled");
            options.cancel ? that.controls.cancel.removeClass("disabled ms-Button--primary").addClass("ms-Button--primary") : that.controls.cancel.removeClass("ms-Button--primary").addClass("disabled");
        },
        select: function () {
            that.keyword = "";
            that.controls.keyword.val("");
            that.ui.dft();
        },
        reset: function () {
            that.ui.clear();
            that.ui.dft();
            that.ui.status({ next: false });
            that.controls.sourceTypeNav.removeClass("is-selected");
            that.controls.main.removeClass("step-second").addClass("step-first");
            that.controls.stepFirstMain.removeClass("selected-file");
        },
        sources: function (options) {
            var _f = false, _dt = options.data != null ? options.data.SourcePoints : [], _d = [], _da = [], _dl = [];
            var _st = that.utility.sortTypeAdd();
            // counter for each source type
            var _ia = 0, _ib = 0, _ic = 0;
            that.ui.status({ next: false });
            that.controls.resultList.scrollTop(0);
            that.controls.resultList.html("");
            _d = _dt;

            // sort by selected sort type
            if (_st.sortType == app.sortTypes.name) {
                if (_st.sortOrder == app.sortOrder.asc) {
                    _d.sort(function (_a, _b) {
                        return (app.string(_a.Name).toUpperCase() > app.string(_b.Name).toUpperCase()) ? 1 : (app.string(_a.Name).toUpperCase() < app.string(_b.Name).toUpperCase()) ? -1 : 0;
                    });
                }
                else {
                    _d.sort(function (_a, _b) {
                        return (app.string(_a.Name).toUpperCase() < app.string(_b.Name).toUpperCase()) ? 1 : (app.string(_a.Name).toUpperCase() > app.string(_b.Name).toUpperCase()) ? -1 : 0;
                    });
                }
            }
            else if (_st.sortType == app.sortTypes.value) {
                if (_st.sortOrder == app.sortOrder.asc) {
                    _d.sort(function (_a, _b) {
                        return (_a.Value.toUpperCase() > _b.Value.toUpperCase())
                            ? 1 : (_a.Value.toUpperCase() < _b.Value.toUpperCase()) ? -1 : 0;
                    });
                }
                else {
                    _d.sort(function (_a, _b) {
                        return (_a.Value.toUpperCase() < _b.Value.toUpperCase())
                            ? 1 : (_a.Value.toUpperCase() > _b.Value.toUpperCase()) ? -1 : 0;
                    });
                }
            }
            else if (_st.sortType == app.sortTypes.type) {
                if (_st.sortOrder == app.sortOrder.asc) {
                    _d.sort(function (_a, _b) {
                        return (_a.SourceType > _b.SourceType) ? 1 : (_a.SourceType < _b.SourceType) ? -1 : 0;
                    });
                }
                else {
                    _d.sort(function (_a, _b) {
                        return (_a.SourceType < _b.SourceType) ? 1 : (_a.SourceType > _b.SourceType) ? -1 : 0;
                    });
                }
            }

            if (options.keyword != undefined && $.trim(options.keyword) != "") {
                var _sk = app.search.splitKeyword({ keyword: $.trim(options.keyword) });
                if (_sk.length > 26) {
                    that.popup.message({ success: false, title: "Only support less then 26 keywords." });
                }
                else {
                    $.each(_d, function (i, d) {
                        var _wi = app.search.weight({ keyword: _sk, source: d });
                        if (_wi > 0) {
                            _da.push(d);
                            _f = true;
                        }
                    });
                    $.each(_da, function (i, d) {
                        var _wi = app.search.weight({ keyword: _sk, source: d });
                        if (_wi > 0) {
                            if (d.SourceType == app.sourceTypes.point) {
                                _ia++;
                            } else if (d.SourceType == app.sourceTypes.chart) {
                                _ib++;
                            } else {
                                _ic++;
                            }
                        }
                    });
                }
            }
            else {
                $.each(_d, function (i, d) {
                    _da.push(d);
                    _f = true;
                });
                $.each(_da, function (i, d) {
                    if (d.SourceType == app.sourceTypes.point) {
                        _ia++;
                    } else if (d.SourceType == app.sourceTypes.chart) {
                        _ib++;
                    } else {
                        _ic++;
                    }
                });
            }

            if (options.sourceType == app.sourceTypes.all) {
                that.controls.sourceTypeNav.removeClass("is-selected");
                if (_ia > 0) {
                    // points
                    $(that.controls.sourceTypeNav[0]).addClass("is-selected");
                    $.each(_da, function (i, d) {
                        if (d.SourceType == app.sourceTypes.point) {
                            _dl.push(d);
                        }
                    });
                } else if (_ib > 0) {
                    // charts
                    $(that.controls.sourceTypeNav[1]).addClass("is-selected");
                    $.each(_da, function (i, d) {
                        if (d.SourceType == app.sourceTypes.chart) {
                            _dl.push(d);
                        }
                    });
                } else if (_ic > 0) {
                    // tables
                    $(that.controls.sourceTypeNav[2]).addClass("is-selected");
                    $.each(_da, function (i, d) {
                        if (d.SourceType == app.sourceTypes.table) {
                            _dl.push(d);
                        }
                    });
                }
            }
            else {
                $.each(_da, function (_a, _b) {
                    if (_b.SourceType == options.sourceType) {
                        _dl.push(_b);
                    }
                });
            }

            that.controls.sourceTypeNav[0].children[1].innerText = _ia;
            that.controls.sourceTypeNav[1].children[1].innerText = _ib;
            that.controls.sourceTypeNav[2].children[1].innerText = _ic;

            $.each(_dl, function (i, d) {
                var _p = that.utility.position(d.Position), _v = d.SourceType == app.sourceTypes.point ? (d.Value ? d.Value : "") : "",
                    _t = d.SourceType == app.sourceTypes.point ? "Source point" : (d.SourceType == app.sourceTypes.table ? "Source table" : "Source chart");

                var _h = '<li class="point-item" data-id="' + d.Id + '" data-file="' + that.utility.fileName(options.data.Name) + '" data-name="' + d.Name + '" title="' + _t + '">';
                _h += '<div class="point-item-line">';

                //_h += '<div class="i1"><div class="ckb-wrapper"><input type="checkbox"/><i class="ms-Icon ms-Icon--CheckMark"></i></div></div>';

                _h += '<div class="i2"><span class="s-name" title="' + d.Name + '">' + d.Name + ' | <i>' + _p.sheet + ' [' + _p.cell + ']</i>' + '</span>';
                _h += '</div>';

                _h += '<div class="i3" title="' + _v + '">' + _v + '</div>';

                _h += '</div>';

                /* Edit format */
                _h += '<div class="add-point-customformat">';
                if (d.SourceType == app.sourceTypes.point) {
                    _h += '<span class="i-preview">Preview: <strong class="lbPreviewValue"></strong></span>';
                }
                _h += '<div class="addCustomFormat add-point-format">';
                if (d.SourceType == app.sourceTypes.point) {
                    _h += '<div class="add-point-place">';
                    _h += '<span class="i-decimal">Decimal place: ';
                    _h += '<i class="i-increase" title="Increase decimal places"></i><i class="i-decrease" title="Decrease decimal places"></i></span>';
                    _h += '</div>';
                }
                _h += '<div class="add-point-box">';
                if (d.SourceType == app.sourceTypes.point) {
                    _h += '<div class="add-point-select">';
                    _h += '<a class="btnSelectFormat" href="javascript:"></a>';
                    _h += '<i class="iconSelectFormat ms-Icon ms-Icon--ChevronDown"></i>';
                    _h += '<ul class="listFormats">';
                    _h += '</ul>';
                    _h += '</div>';
                }
                _h += '<button class="ms-Button ms-Button--small ms-Button--primary i-add">';
                _h += '<span class="ms-Button-label">' + (d.SourceType == app.sourceTypes.table ? "Add as Table" : "Add") + '</span>';
                _h += '</button>';
                if (d.SourceType == app.sourceTypes.table) {
                    _h += '<button class="ms-Button ms-Button--small ms-Button--primary i-add i-table-image">';
                    _h += '<span class="ms-Button-label">Add as Image</span>';
                    _h += '</button>';
                }
                _h += '</div>';
                _h += '</div>';
                _h += '<div class="clear"></div>';
                _h += '</div>';
                /* Edit format end */

                _h += '</li>';

                $(_h).appendTo(that.controls.resultList);
            });

            _f ? that.controls.resultNotFound.hide() : that.controls.resultNotFound.show();
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
                            if (app.search.weight({ keyword: _sk, source: d.ReferencedSourcePoint }) > 0) {
                                result.push(d);
                            }
                        });
                    }
                }
                else {
                    result = original;
                }
                return result;
            },
            filterError: function (original) {
                return that.errorFilter.filterListByError(original);
            }
        },
        list: function (options, callback) {
            that.range.all(function (result) {
                var _tags = [];
                if (result.status == app.status.succeeded) {
                    _tags = result.data;
                }
                that.range.values({ data: $.extend([], that.points), tags: _tags, refresh: options.refresh }, function (result) {

                    var _dt = result.data, _d = [], _ss = [], _st = that.utility.sortType(), _ns = false;
                    var _pt = that.utility.selectedSourceType(".point-types-mana");
                    $.each(_dt, function (_io, _ie) {
                        if (_ie.changed) {
                            _ns = true;
                            return false;
                        }
                    });
                    // records that are filtered by source type and searching keyword 
                    var _df = [];
                    // counter for each source type
                    var _ia = 0, _ib = 0, _ic = 0;
                    // filtered by keyword
                    _d = that.ui.listFilter.search(_dt);
                    //filtered by error
                    _d = that.ui.listFilter.filterError(_d);
                    // filtered by source type
                    if (_pt == app.sourceTypes.all) {
                        _df = _d;
                    }
                    else {
                        $.each(_d, function (_a, _b) {
                            if (_b.ReferencedSourcePoint.SourceType == _pt) {
                                _df.push(_b);
                            }
                        });
                    }

                    $.each(_d, function (i, d) {
                        if (d.ReferencedSourcePoint.SourceType == app.sourceTypes.point) {
                            _ia++;
                        } else if (d.ReferencedSourcePoint.SourceType == app.sourceTypes.chart) {
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
                                return (app.string(_a.ReferencedSourcePoint.Name).toUpperCase() > app.string(_b.ReferencedSourcePoint.Name).toUpperCase()) ? 1 : (app.string(_a.ReferencedSourcePoint.Name).toUpperCase() < app.string(_b.ReferencedSourcePoint.Name).toUpperCase()) ? -1 : 0;
                            });
                        }
                        else {
                            _df.sort(function (_a, _b) {
                                return (app.string(_a.ReferencedSourcePoint.Name).toUpperCase() < app.string(_b.ReferencedSourcePoint.Name).toUpperCase()) ? 1 : (app.string(_a.ReferencedSourcePoint.Name).toUpperCase() > app.string(_b.ReferencedSourcePoint.Name).toUpperCase()) ? -1 : 0;
                            });
                        }
                    }
                    else if (_st.sortType == app.sortTypes.value) {
                        if (_pt == app.sourceTypes.point) {
                            if (_st.sortOrder == app.sortOrder.asc) {
                                _df.sort(function (_a, _b) {
                                    return (_a.ReferencedSourcePoint.Value.toUpperCase() > _b.ReferencedSourcePoint.Value.toUpperCase())
                                        ? 1 : (_a.ReferencedSourcePoint.Value.toUpperCase() < _b.ReferencedSourcePoint.Value.toUpperCase()) ? -1 : 0;
                                });
                            }
                            else {
                                _df.sort(function (_a, _b) {
                                    return (_a.ReferencedSourcePoint.Value.toUpperCase() < _b.ReferencedSourcePoint.Value.toUpperCase())
                                        ? 1 : (_a.ReferencedSourcePoint.Value.toUpperCase() > _b.ReferencedSourcePoint.Value.toUpperCase()) ? -1 : 0;
                                });
                            }
                        }
                        else {
                            if (_st.sortOrder == app.sortOrder.asc) {
                                _df.sort(function (_a, _b) {
                                    return (_a.PublishedStatus > _b.PublishedStatus)
                                        ? 1 : (_a.PublishedStatus < _b.PublishedStatus) ? -1 : 0;
                                });
                            }
                            else {
                                _df.sort(function (_a, _b) {
                                    return (_a.PublishedStatus < _b.PublishedStatus)
                                        ? 1 : (_a.PublishedStatus > _b.PublishedStatus) ? -1 : 0;
                                });
                            }
                        }
                    }
                    else if (_st.sortType == app.sortTypes.type) {
                        if (_st.sortOrder == app.sortOrder.asc) {
                            _df.sort(function (_a, _b) {
                                return (_a.ReferencedSourcePoint.SourceType > _b.ReferencedSourcePoint.SourceType) ? 1 : (_a.ReferencedSourcePoint.SourceType < _b.ReferencedSourcePoint.SourceType) ? -1 : 0;
                            });
                        }
                        else {
                            _df.sort(function (_a, _b) {
                                return (_a.ReferencedSourcePoint.SourceType < _b.ReferencedSourcePoint.SourceType) ? 1 : (_a.ReferencedSourcePoint.SourceType > _b.ReferencedSourcePoint.SourceType) ? -1 : 0;
                            });
                        }
                    }
                    else {
                        _df.sort(function (_a, _b) {
                            return (_a.orderBy > _b.orderBy) ? 1 : (_a.orderBy < _b.orderBy) ? -1 : 0;
                        });
                    }

                    if (options.refresh) {
                        var _c = that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked"), _s = that.utility.selected();
                        $.each(_s, function (m, n) {
                            _ss.push(n.DestinationPointId);
                        });
                        that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", _c);
                        if (_c) {
                            that.controls.headerListPoints.find(".point-header .ckb-wrapper").addClass("checked");
                        }
                        else {
                            that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
                        }
                    }
                    else {
                        that.controls.headerListPoints.find(".point-header .ckb-wrapper input").prop("checked", false);
                        that.controls.headerListPoints.find(".point-header .ckb-wrapper").removeClass("checked");
                    }
                    that.utility.scrollTop();
                    that.controls.list.find(".point-item").remove();
                    that.ui.item({ index: 0, data: _df, refresh: options.refresh, selected: _ss, notSync: _ns }, callback);
                    if (that.controls.list.find(">li").length > 0) {
                        that.controls.moveUp.removeClass("disabled");
                        that.controls.moveDown.removeClass("disabled");
                    }
                });
            });
        },
        item: function (options, callback) {
            if (options.index < options.data.length) {
                var _dsp = options.data[options.index], _item = _dsp.ReferencedSourcePoint, _sourceType = _item.SourceType,
                    _sourcePointCatalog = _item.Catalog,
                    _tp = (_dsp.existed == undefined || _dsp.existed == "undefined") ? false : _dsp.existed,
                    _s = _tp && _item.Status == 0;
                if (options.index >= that.pagerSize * (that.pagerIndex - 1) && options.index < that.pagerSize * that.pagerIndex) {
                    var _spv = _item.Value && _item.Value != null ? _item.Value : "",
                        _v = _sourceType == app.sourceTypes.point ? _spv : "",
                        _p = that.utility.position(_item.Position),
                        _fn = that.utility.fileName(_sourcePointCatalog.Name),
                        _sel = $.inArray(_dsp.Id, options.selected) > -1,
                        _fv = _sourceType == app.sourceTypes.point ? that.format.convert({ value: _spv, formats: _dsp.CustomFormats, decimal: _dsp.DecimalPlace }) : _v,
                        _ff = [],
                        _pht = that.utility.publishHistory({ data: _item.PublishedHistories }),
                        _cs = !_dsp.PublishedStatus ? "status-notpublished" : "status-published",
                        _ts = _sourceType == app.sourceTypes.point ? "type-point" : (_sourceType == app.sourceTypes.table ? "type-table" : "type-chart"),
                        _fd = !_sourcePointCatalog.IsDeleted;
                    $.each(_dsp.CustomFormats != null ? _dsp.CustomFormats : [], function (_x, _y) { _ff.push(_y.DisplayName); });
                    if (_dsp.DecimalPlace != null) {
                        _ff.push("Displayed decimals");
                    }
                    var _cf = _ff.join("; ").replace(/"/g, "&quot;");
                    var _h = '<li class="point-item' + (_s && _fd ? "" : " item-error") + ' ' + _cs + ' ' + _ts + '" data-id="' + _dsp.Id + '" data-value="' + _v + '" data-range="' + _dsp.RangeId + '" data-sourcetype="' + _sourceType + '">';
                    _h += '<div class="point-item-line">';
                    // i1
                    _h += '<div class="i1"><div class="ckb-wrapper' + (_sel ? " checked" : "") + '"><input type="checkbox" ' + (_sel ? 'checked="checked"' : '') + ' /><i class="ms-Icon ms-Icon--CheckMark"></i></div></div>';
                    // i2
                    _h += '<div class="i2"><span class="s-name" title="' + _item.Name + '">' + _item.Name + '</span>';
                    _h += '<div class="sp-file-pos">';
                    if (_fd) {
                        _h += '<span class="i-file" title="' + _sourcePointCatalog.Name + '" data-path="' + _sourcePointCatalog.Name + '">' + _fn + ' </span>';
                    }
                    _h += '<span title="' + (_p.sheet ? _p.sheet : "") + ':[' + (_p.cell ? _p.cell : "") + ']">/ ' + (_p.sheet ? _p.sheet : "") + ':[' + (_p.cell ? _p.cell : "") + ']</span>';
                    _h += '</div>';
                    if (_sourceType == app.sourceTypes.point) {
                        _h += '<span class="item-formatted" title="' + _fv + '">' + (_ff.length > 0 ? "Formatted value" : "Source point value") + ': <strong>' + _fv + '</strong></span>';
                        _h += '<span class="item-formats" title="' + (_ff.length > 0 ? _cf : "No custom formatting applied") + '">Format: <strong>' + (_ff.length > 0 ? _cf : "No custom formatting applied") + '</strong></span>';
                    }
                    else {
                        _h += '<span class="item-formatted">&nbsp;</span>';
                        _h += '<span class="item-formats">&nbsp;</span>';
                    }

                    _h += '</div>';
                    // i3
                    if (_sourceType == app.sourceTypes.point) {
                        _h += '<div class="i3" title="' + _spv + '">' + _spv + '</div>';
                    }
                    else {
                        _h += '<div class="i3"><i class="ms-Icon ms-Icon--Error"></i><i class="ms-Icon ms-Icon--Completed"></i></div>';
                    }
                    // i5
                    _h += '<div class="i5">';
                    _h += '<div class="i-menu"><a href="javascript:"><span title="Action"><i class="ms-Icon ms-Icon--More"></i></span><span class="quick-menu"><span class="i-history" title="History"><i class="ms-Icon ms-Icon--ChevronRight"></i><i>History</i></span><span class="i-delete" title="Delete"><i class="ms-Icon ms-Icon--Cancel"></i><i>Delete</i></span><span class="i-edit" title="Edit Custom Format"><i class="ms-Icon ms-Icon--Edit"></i><i>Edit</i><span></span></a></div>';
                    _h += '</div>';
                    // 
                    _h += '<div class="clear"></div>';
                    _h += '</div>';

                    /* Edit format */
                    _h += '<div class="add-point-customformat">';
                    if (_sourceType == app.sourceTypes.point) {
                        _h += '<span class="i-preview">Preview: <strong class="lbPreviewValue"></strong></span>';
                    }
                    _h += '<div class="addCustomFormat add-point-format">';
                    if (_sourceType == app.sourceTypes.point) {
                        _h += '<div class="add-point-place">';
                        _h += '<span class="i-decimal">Decimal place: ';
                        _h += '<i class="i-increase" title="Increase decimal places"></i><i class="i-decrease" title="Decrease decimal places"></i></span>';
                        _h += '</div>';
                    }
                    _h += '<div class="add-point-box">';
                    if (_sourceType == app.sourceTypes.point) {
                        _h += '<div class="add-point-select">';
                        _h += '<a class="btnSelectFormat" href="javascript:"></a>';
                        _h += '<i class="iconSelectFormat ms-Icon ms-Icon--ChevronDown"></i>';
                        _h += '<ul class="listFormats">';
                        _h += '</ul>';
                        _h += '</div>';
                    }
                    _h += '<button class="ms-Button ms-Button--small ms-Button--primary i-update">';
                    _h += '<span class="ms-Button-label">Update</span>';
                    _h += '</button>';
                    _h += '</div>';
                    _h += '</div>';
                    _h += '<div class="clear"></div>';
                    _h += '</div>';
                    /* Edit format end */

                    // History
                    _h += '<div class="item-history"><ul class="history-list">';
                    _h += '<li class="history-header"><div class="h0"></div><div class="h1"><span>Name</span></div><div class="h2"><span>Date Modified</span></div><div class="h3"><span>Value</span></div></li>';
                    $.each(_pht, function (m, n) {
                        var __cv = (n.Value ? n.Value : ""), __pvr = _sourceType == app.sourceTypes.point ? __cv : (__cv == "Cloned" ? "Cloned" : "Published");
                        _h += '<li class="history-item"><div class="h0"></div><div class="h1" title="' + n.PublishedUser + '">' + n.PublishedUser +
                            '</div><div class="h2" title="' + app.date(n.PublishedDate) + '">' + app.date(n.PublishedDate) +
                            '</div><div class="h3" title="' + __pvr + '">' + __pvr + '</div></li>';
                    });
                    _h += '</ul>';
                    _h += '<div class="clear"></div>';
                    _h += '</div>';

                    // Error
                    _h += '<div class="error-info">';
                    _h += '<div class="e1"><i class="ms-Icon ms-Icon--ErrorBadge"></i></div>';
                    _h += '<div class="e2">';
                    if (!_fd) {
                        _h += '<p><strong>Error</strong>:The source file used for this destination point has been deleted from the catalog, please delete the destination point and use a valid source point.</p>';
                    }
                    else {
                        if (!_tp) {
                            _h += '<p><strong>Error</strong>:The content control that was bound to the source point for this destination point has been removed, please deleted destination point.</p>';
                        }
                        else if (_item.Status == 1) {
                            _h += '<p><strong>Error</strong>:The source point used for this destination point has been deleted from the catalog, please delete the destination point and use a valid source point.</p>';
                        }
                    }
                    _h += '</div>';
                    _h += '<div class="clear"></div>';
                    _h += '</div>';

                    _h += '</li>';
                    that.controls.list.append(_h);
                }

                options.index++;
                that.ui.item(options, callback);
            }
            else {
                if (that.firstLoad && options.notSync) {
                    that.popup.message({ success: false, title: "The document is not in sync with the catalog, please click refresh to synchronize.", canClose: true }, function () { });
                }
                that.firstLoad = false;
                if (callback) {
                    callback({ status: app.status.succeeded });
                }
            }
        },
        remove: function (options) {
            that.controls.list.find("[data-id=" + options.Id + "]").remove();
        },
        customFormat: function (options, callback) {

            var _currentItem = options.o.hasClass("point-item") ? options.o : options.o.closest(".ponit-item");

            _currentItem.find(".listFormats").html("");
            var _si = [], _sn = [], _sd = [];
            if (options.selected && options.selected != null) {
                $.each(options.selected.CustomFormats, function (_x, _y) {
                    _si.push(_y.Id);
                    _sn.push(_y.Name);
                    _sd.push(_y.DisplayName);
                });
                _currentItem.find(".btnSelectFormat").html(_sd.length > 0 ? _sd.join(", ") : "None");
                _currentItem.find(".btnSelectFormat").prop("title", _sd.length > 0 ? _sd.join(", ") : "None");
                _currentItem.find(".btnSelectFormat").prop("selected", _si.join(","));
                _currentItem.find(".btnSelectFormat").prop("name", _sn.join(","));
                _currentItem.find(".btnSelectFormat").prop("place", options.selected.DecimalPlace && options.selected.DecimalPlace != null ? options.selected.DecimalPlace : "");
            }
            else {
                _currentItem.find(".btnSelectFormat").html("None");
                _currentItem.find(".btnSelectFormat").prop("title", "None");
                _currentItem.find(".btnSelectFormat").prop("selected", "");
                _currentItem.find(".btnSelectFormat").prop("name", "");
                _currentItem.find(".btnSelectFormat").prop("place", "");
            }

            var _v = options.ref ? options.selected.ReferencedSourcePoint.Value : options.selectedPoint.Value; //_dataList.find("li." + _itemSelectedName + " .btnSelectFormat").prop("original");
            _currentItem.find(".listFormats").removeClass("convert1 convert2 convert3 convert4");
            _currentItem.find(".lbPreviewValue").html(_v);
            _currentItem.find(".addCustomFormat").removeClass("selected-number selected-date");
            if (that.format.isNumber(_v)) {
                _currentItem.find(".addCustomFormat").addClass("selected-number");
            }
            else if (that.format.isDate(_v)) {
                _currentItem.find(".addCustomFormat").addClass("selected-date");
            }

            if (_sn.length > 0) {
                var _tn = $.inArray("ConvertToThousands", _sn) > -1 ? "IncludeThousandDescriptor" : ($.inArray("ConvertToMillions", _sn) > -1 ? "IncludeMillionDescriptor" : ($.inArray("ConvertToBillions", _sn) > -1 ? "IncludeBillionDescriptor" : ($.inArray("ConvertToHundreds", _sn) > -1 ? "IncludeHundredDescriptor" : "")));
                var _cl = $.inArray("ConvertToThousands", _sn) > -1 ? "convert2" : ($.inArray("ConvertToMillions", _sn) > -1 ? "convert3" : ($.inArray("ConvertToBillions", _sn) > -1 ? "convert4" : ($.inArray("ConvertToHundreds", _sn) > -1 ? "convert1" : "")));
                _currentItem.find(".listFormats").addClass(_cl);
                _currentItem.find(".listFormats").find("ul > li[data-name=" + _tn + "]").addClass("checked");
            }

            var _dd = [], _dt = [];
            $.each(options.data, function (_i, _e) {
                var _i = $.inArray(_e.GroupName, _dt);
                if (_i == -1) {
                    _dd.push({ Name: _e.GroupName, OrderBy: _e.GroupOrderBy, Formats: [{ Id: _e.Id, Name: _e.Name, DisplayName: _e.DisplayName, Description: _e.Description, OrderBy: _e.OrderBy }] });
                    _dt.push(_e.GroupName);
                }
                else {
                    _dd[_i].Formats.push({ Id: _e.Id, Name: _e.Name, DisplayName: _e.DisplayName, Description: _e.Description, OrderBy: _e.OrderBy });
                }
            });
            $.each(_dd, function (_i, _e) {
                _e.Formats.sort(function (_m, _n) {
                    return _m.OrderBy > _n.OrderBy ? 1 : _m.OrderBy < _n.OrderBy ? -1 : 0;
                });
            });
            _dd.sort(function (_m, _n) {
                return _m.OrderBy > _n.OrderBy ? 1 : _m.OrderBy < _n.OrderBy ? -1 : 0;
            });

            if (_dd) {
                $.each(_dd, function (m, n) {
                    var _h = '', _c = '';
                    _c = (n.Name == "Convert to" || n.Name == "Negative number" || n.Name == "Descriptor") ? "value-number" : (n.Name == "Symbol") ? "value-string" : "value-date";
                    _h += '<li class="' + (n.Name == "Descriptor" ? "drp-checkbox drp-descriptor " : "drp-radio ") + '' + _c + '">';
                    if (n.Name != "Descriptor") {
                        _h += '<label>' + n.Name + '</label>';
                    }
                    _h += '<ul>';
                    $.each(n.Formats, function (i, d) {
                        _h += '<li data-id="' + d.Id + '" data-name="' + d.Name + '" title="' + d.Description + '" class="' + ($.inArray(d.Id, _si) > -1 ? "checked" : "") + '" data-displayname="' + d.DisplayName.replace(/"/g, "&quot;") + '">';
                        _h += '<div><i></i></div>';
                        _h += '<a href="javascript:">' + (n.Name == "Descriptor" ? "Descriptor" : d.DisplayName) + '</a>';
                        _h += '</li>';
                    });
                    _h += '</ul>';
                    _h += '</li>';
                    _currentItem.find(".listFormats").append(_h);
                });
            }
            if (callback) {
                callback();
            }
        },
        recentFiles: function (options, callback) {
            that.controls.recentFiles.html("");
            if (options.data.length > 0) {
                $.each(options.data, function (m, n) {
                    var _h = '<li class="ms-Dropdown-item">' + that.utility.fileName(n.Name) + '</li>';
                    that.controls.recentFiles.append(_h);
                });
            }
            else {
                that.controls.recentFiles.html("No recent file.");
            }
            if (callback) {
                callback();
            }
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
        }
    };

    return that;
})();