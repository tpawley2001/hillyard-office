/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of Hillyard Office.
 *
 * Copyright (c) 2026 Hillyard Tech
 * SPDX-License-Identifier: MIT
 */

#include <notebookbar/Backstage.hxx>

#include <recentdocsview.hxx>
#include <templatedefaultview.hxx>
#include <sfx2/templatelocalview.hxx>
#include <sfx2/sfxresid.hxx>
#include <sfx2/strings.hrc>

#include <comphelper/processfactory.hxx>
#include <comphelper/propertyvalue.hxx>
#include <comphelper/dispatchcommand.hxx>
#include <config_folders.h>
#include <osl/file.hxx>
#include <rtl/bootstrap.hxx>
#include <svl/numformat.hxx>
#include <tools/datetime.hxx>
#include <tools/urlobj.hxx>
#include <unotools/configmgr.hxx>
#include <unotools/localedatawrapper.hxx>
#include <unotools/syslocale.hxx>
#include <vcl/commandinfoprovider.hxx>
#include <vcl/event.hxx>
#include <vcl/notebookbar/notebookbar.hxx>
#include <vcl/print.hxx>
#include <vcl/syswin.hxx>
#include <vcl/settings.hxx>
#include <vcl/svapp.hxx>
#include <vcl/vclevent.hxx>
#include <vcl/virdev.hxx>
#include <vcl/weld/Builder.hxx>
#include <vcl/weld/ScrolledWindow.hxx>
#include <vcl/weld/Menu.hxx>
#include <toolkit/helper/vclunohelper.hxx>

#include <com/sun/star/beans/XPropertySet.hpp>
#include <com/sun/star/container/XIndexAccess.hpp>
#include <com/sun/star/document/MacroExecMode.hpp>
#include <com/sun/star/document/UpdateDocMode.hpp>
#include <com/sun/star/document/XDocumentProperties.hpp>
#include <com/sun/star/document/XDocumentPropertiesSupplier.hpp>
#include <com/sun/star/drawing/XDrawPagesSupplier.hpp>
#include <com/sun/star/frame/Desktop.hpp>
#include <com/sun/star/frame/ModuleManager.hpp>
#include <com/sun/star/frame/XComponentLoader.hpp>
#include <com/sun/star/frame/XController.hpp>
#include <com/sun/star/frame/XModel.hpp>
#include <com/sun/star/frame/XTitle.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/task/InteractionHandler.hpp>
#include <com/sun/star/util/DateTime.hpp>

using namespace css;

namespace
{
struct BackstageCommand
{
    OUString aId;
    OUString aCommand;
    OUString aIcon;
};

const BackstageCommand aCommands[] = {
#include "BackstageCommands.inc"
};

OUString CommandFor(std::u16string_view rId)
{
    for (const auto& r : aCommands)
        if (r.aId == rId)
            return r.aCommand;
    return OUString();
}

constexpr OUString aPages[]
    = { u"home"_ustr,  u"new"_ustr,   u"open"_ustr,   u"info"_ustr,   u"saveas"_ustr,
        u"print"_ustr, u"share"_ustr, u"export"_ustr, u"account"_ustr };

constexpr OUString aRail[] = { u"home"_ustr,   u"new"_ustr,    u"open"_ustr,  u"info"_ustr,
                               u"save"_ustr,   u"saveas"_ustr, u"print"_ustr, u"share"_ustr,
                               u"export"_ustr, u"close"_ustr,  u"account"_ustr,
                               u"options"_ustr };

struct PendingCommand
{
    uno::Reference<frame::XFrame> xFrame;
    OUString aCommand;
};

void DispatchPending(void*, void* p)
{
    std::unique_ptr<PendingCommand> xPending(static_cast<PendingCommand*>(p));
    try
    {
        comphelper::dispatchCommand(xPending->aCommand, xPending->xFrame, {});
    }
    catch (const uno::Exception&)
    {
    }
}

bool IsPage(std::u16string_view rId)
{
    return std::find(std::begin(aPages), std::end(aPages), rId) != std::end(aPages);
}

constexpr tools::Long RAIL_WIDTH = 200;
constexpr tools::Long RAIL_ITEM_HEIGHT = 40;

// Each Office app's Backstage colour.
Color AccentFor(std::u16string_view rModule)
{
    if (rModule == u"com.sun.star.sheet.SpreadsheetDocument")
        return Color(0x10, 0x7C, 0x41); // Excel
    if (rModule == u"com.sun.star.presentation.PresentationDocument")
        return Color(0xC4, 0x3E, 0x1C); // PowerPoint
    if (rModule == u"com.sun.star.drawing.DrawingDocument")
        return Color(0x39, 0x55, 0xA3); // Visio
    return Color(0x18, 0x5A, 0xBD); // Word
}

std::map<vcl::Window*, VclPtr<Backstage>>& Instances()
{
    static std::map<vcl::Window*, VclPtr<Backstage>> aInstances;
    return aInstances;
}

OUString FormatDate(const util::DateTime& rDT)
{
    if (rDT.Year == 0)
        return OUString();
    const LocaleDataWrapper& rLocale = SvtSysLocale().GetLocaleData();
    DateTime aDT(rDT);
    return rLocale.getDate(aDT) + " " + rLocale.getTime(aDT, false);
}

OUString FormatSize(sal_uInt64 nBytes)
{
    const LocaleDataWrapper& rLocale = SvtSysLocale().GetLocaleData();
    if (nBytes < 1024)
        return OUString::number(nBytes) + " bytes";
    if (nBytes < 1024 * 1024)
        return rLocale.getNum(nBytes * 10 / 1024, 1) + "KB";
    return rLocale.getNum(nBytes * 10 / (1024 * 1024), 1) + "MB";
}
}

Backstage::Backstage(vcl::Window* pParent, uno::Reference<frame::XFrame> xFrame)
    : InterimItemWindow(pParent, u"sfx/ui/backstage.ui"_ustr, u"Backstage"_ustr, false)
    , mxFrame(std::move(xFrame))
    , mbTemplatesFilled(false)
    , mxRail(m_xBuilder->weld_container(u"rail"_ustr))
    , mxBack(m_xBuilder->weld_button(u"back"_ustr))
    , mxPages(m_xBuilder->weld_notebook(u"pages"_ustr))
    , mxGreeting(m_xBuilder->weld_label(u"greeting"_ustr))
    , mxHomeBlank(m_xBuilder->weld_button(u"home_blank"_ustr))
    , mxHomeMore(m_xBuilder->weld_button(u"home_more"_ustr))
    , mxNewBlank(m_xBuilder->weld_button(u"new_blank"_ustr))
    , mxHomeRecent(new sfx2::RecentDocsView(
          m_xBuilder->weld_scrolled_window(u"scroll_home_recent"_ustr, true)))
    , mxHomeRecentWin(new weld::CustomWeld(*m_xBuilder, u"home_recent"_ustr, *mxHomeRecent))
    , mxOpenRecent(new sfx2::RecentDocsView(
          m_xBuilder->weld_scrolled_window(u"scroll_open_recent"_ustr, true)))
    , mxOpenRecentWin(new weld::CustomWeld(*m_xBuilder, u"open_recent"_ustr, *mxOpenRecent))
    , mxTemplates(new TemplateDefaultView(m_xBuilder->weld_scrolled_window(u"scroll_templates"_ustr, true),
                                          m_xBuilder->weld_menu(u"templatemenu"_ustr)))
    , mxTemplatesWin(new weld::CustomWeld(*m_xBuilder, u"templates"_ustr, *mxTemplates))
    , mxInfoName(m_xBuilder->weld_label(u"info_name"_ustr))
    , mxInfoPath(m_xBuilder->weld_label(u"info_path"_ustr))
    , mxPrinter(m_xBuilder->weld_label(u"print_printer"_ustr))
    , mxProduct(m_xBuilder->weld_label(u"account_product"_ustr))
    , mxVersion(m_xBuilder->weld_label(u"account_version"_ustr))
{
    SetPaintTransparent(false);
    SetStyle(GetStyle() | WB_DIALOGCONTROL);

    try
    {
        maModule = frame::ModuleManager::create(comphelper::getProcessComponentContext())
                       ->identify(mxFrame);
    }
    catch (const uno::Exception&)
    {
    }
    maAccent = AccentFor(maModule);

    mxBack->connect_clicked(LINK(this, Backstage, BackHdl));
    for (const OUString& rId : aRail)
    {
        auto xButton = m_xBuilder->weld_button("rail_" + rId);
        xButton->connect_clicked(LINK(this, Backstage, RailHdl));
        xButton->set_size_request(RAIL_WIDTH, RAIL_ITEM_HEIGHT);
        maRail.emplace_back(rId, std::move(xButton));
    }
    mxBack->set_size_request(RAIL_WIDTH, RAIL_ITEM_HEIGHT + 16);

    for (const auto& r : aCommands)
    {
        if (r.aId.startsWith("rail_"))
            continue;
        auto xButton = m_xBuilder->weld_button(r.aId);
        xButton->connect_clicked(LINK(this, Backstage, ActionHdl));
        xButton->set_image(
            vcl::CommandInfoProvider::GetXGraphicForCommand(r.aIcon, mxFrame, vcl::ImageType::Size32));
        maActions.emplace(r.aId, std::move(xButton));
    }
    // Writer alone exports EPUB.
    if (maModule != u"com.sun.star.text.TextDocument")
    {
        maActions[u"export_epub"_ustr]->hide();
        m_xBuilder->weld_widget(u"export_epub_title"_ustr)->hide();
        m_xBuilder->weld_widget(u"export_epub_desc"_ustr)->hide();
    }

    for (const OUString& rPage : aPages)
        if (auto xTitle = m_xBuilder->weld_label("title_" + rPage))
            maTitles.push_back(std::move(xTitle));

    for (const char* pKey : { "size", "pages", "words", "title", "author", "modified", "created" })
    {
        OUString aKey = OUString::createFromAscii(pKey);
        maProps.emplace(aKey, m_xBuilder->weld_label("prop_" + aKey));
    }

    // Blank: the current application's new document, as in Office.
    auto xBlank = vcl::CommandInfoProvider::GetXGraphicForCommand(u".uno:AddDirect"_ustr, mxFrame,
                                                                  vcl::ImageType::Size32);
    for (weld::Button* pButton : { mxHomeBlank.get(), mxNewBlank.get() })
    {
        pButton->set_image(xBlank);
        pButton->connect_clicked(LINK(this, Backstage, BlankHdl));
    }
    OUString aBlank = maModule == u"com.sun.star.sheet.SpreadsheetDocument"
                          ? SfxResId(STR_BACKSTAGE_BLANK_WORKBOOK)
                      : maModule == u"com.sun.star.presentation.PresentationDocument"
                          ? SfxResId(STR_BACKSTAGE_BLANK_PRESENTATION)
                          : SfxResId(STR_BACKSTAGE_BLANK_DOCUMENT);
    mxHomeBlank->set_label(aBlank);
    mxNewBlank->set_label(aBlank);
    mxHomeMore->connect_clicked(LINK(this, Backstage, MoreHdl));

    // Only this application's files, as in Office.
    sfx2::ApplicationType eType = sfx2::ApplicationType::TYPE_WRITER;
    if (maModule == u"com.sun.star.sheet.SpreadsheetDocument")
        eType = sfx2::ApplicationType::TYPE_CALC;
    else if (maModule == u"com.sun.star.presentation.PresentationDocument")
        eType = sfx2::ApplicationType::TYPE_IMPRESS;
    else if (maModule == u"com.sun.star.drawing.DrawingDocument")
        eType = sfx2::ApplicationType::TYPE_DRAW;
    mxHomeRecent->setFilter(eType);
    mxOpenRecent->setFilter(eType);

    mxTemplates->setOpenTemplateHdl(LINK(this, Backstage, OpenTemplateHdl));

    OUString aProduct = utl::ConfigManager::getProductName() + " "
                        + utl::ConfigManager::getProductVersion();
    mxProduct->set_label(aProduct);
    OUString aBuildId(u"${$BRAND_BASE_DIR/" LIBO_ETC_FOLDER "/" SAL_CONFIGFILE("version") ":buildid}"_ustr);
    rtl::Bootstrap::expandMacros(aBuildId);
    mxVersion->set_label(utl::ConfigManager::getAboutBoxProductVersion()
                         + (aBuildId.isEmpty() ? OUString() : " (" + aBuildId.copy(0, std::min<sal_Int32>(9, aBuildId.getLength())) + ")"));

    ApplyColors();

    GetParent()->AddEventListener(LINK(this, Backstage, ParentEventHdl));
    Hide();
}

Backstage::~Backstage() { disposeOnce(); }

void Backstage::dispose()
{
    for (auto it = Instances().begin(); it != Instances().end(); ++it)
        if (it->second.get() == this)
        {
            Instances().erase(it);
            break;
        }
    if (GetParent())
        GetParent()->RemoveEventListener(LINK(this, Backstage, ParentEventHdl));
    mxTemplatesWin.reset();
    mxTemplates.reset();
    mxOpenRecentWin.reset();
    mxOpenRecent.reset();
    mxHomeRecentWin.reset();
    mxHomeRecent.reset();
    maProps.clear();
    maTitles.clear();
    maActions.clear();
    maRail.clear();
    mxInfoName.reset();
    mxInfoPath.reset();
    mxPrinter.reset();
    mxProduct.reset();
    mxVersion.reset();
    mxGreeting.reset();
    mxHomeBlank.reset();
    mxHomeMore.reset();
    mxNewBlank.reset();
    mxPages.reset();
    mxBack.reset();
    mxRail.reset();
    mxFrame.clear();
    InterimItemWindow::dispose();
}

VclPtr<Backstage> Backstage::Get(const uno::Reference<frame::XFrame>& xFrame)
{
    if (!xFrame.is())
        return nullptr;
    VclPtr<vcl::Window> pParent = VCLUnoHelper::GetWindow(xFrame->getContainerWindow());
    if (!pParent)
        return nullptr;
    auto& rInstances = Instances();
    auto it = rInstances.find(pParent.get());
    if (it != rInstances.end() && !it->second->isDisposed())
        return it->second;
    VclPtr<Backstage> xNew = VclPtr<Backstage>::Create(pParent, xFrame);
    rInstances[pParent.get()] = xNew;
    return xNew;
}

void Backstage::FitToParent()
{
    if (vcl::Window* pParent = GetParent())
        SetPosSizePixel(Point(0, 0), pParent->GetOutputSizePixel());
}

IMPL_LINK(Backstage, ParentEventHdl, VclWindowEvent&, rEvent, void)
{
    if (rEvent.GetId() == VclEventId::WindowResize && IsVisible())
        FitToParent();
    else if (rEvent.GetId() == VclEventId::ObjectDying)
        disposeOnce();
}

void Backstage::ApplyColors()
{
    mxRail->set_background(maAccent);
    const StyleSettings& rStyle = Application::GetSettings().GetStyleSettings();
    // Opaque: the document must not show through the pages.
    SetBackground(rStyle.GetWindowColor());
    m_xContainer->set_background(rStyle.GetWindowColor());
    mxPages->set_background(rStyle.GetWindowColor());
    mxHomeRecent->SetBackstageLook();
    mxOpenRecent->SetBackstageLook();
    mxTemplates->SetBackstageLook();

    vcl::Font aTitle(mxInfoName->get_font());
    aTitle.SetFontHeight(aTitle.GetFontHeight() * 2);
    for (auto& rTitle : maTitles)
        rTitle->set_font(aTitle);
    mxGreeting->set_font(aTitle);
    mxInfoName->set_font(aTitle);
    vcl::Font aBig(mxProduct->get_font());
    aBig.SetFontHeight(aBig.GetFontHeight() * 3 / 2);
    mxProduct->set_font(aBig);
    PaintRail();
}

// The rail is the app's colour with white text; the open page is darker.
void Backstage::PaintRail()
{
    vcl::Font aFont(Application::GetSettings().GetStyleSettings().GetAppFont());
    aFont.SetFontSize(Size(0, 16)); // pixels: Office's rail is about 15 px Segoe UI
    aFont.SetColor(COL_WHITE);

    auto paint = [&](weld::Button& rButton, const OUString& rText, bool bSelected, bool bArrow)
    {
        ScopedVclPtrInstance<VirtualDevice> xDev;
        Size aSize(RAIL_WIDTH, bArrow ? RAIL_ITEM_HEIGHT + 16 : RAIL_ITEM_HEIGHT);
        xDev->SetOutputSizePixel(aSize);
        Color aBack(maAccent);
        if (bSelected)
            aBack.DecreaseLuminance(48);
        xDev->SetBackground(Wallpaper(aBack));
        xDev->Erase();
        vcl::Font aItemFont(aFont);
        if (bArrow)
            aItemFont.SetFontSize(Size(0, 26));
        if (bSelected)
            aItemFont.SetWeight(WEIGHT_BOLD);
        xDev->SetFont(aItemFont);
        xDev->SetTextColor(COL_WHITE);
        OUString aText(rText.replaceAll("_", ""));
        tools::Long nY = (aSize.Height() - xDev->GetTextHeight()) / 2;
        xDev->DrawText(Point(24, nY), aText);
        rButton.set_custom_button(xDev.get());
        rButton.set_accessible_name(aText);
    };

    paint(*mxBack, u"←"_ustr, false, true);
    for (auto& [rId, xButton] : maRail)
        paint(*xButton, xButton->get_label(), rId == maPage, false);
}

void Backstage::DataChanged(const DataChangedEvent& rDCEvt)
{
    InterimItemWindow::DataChanged(rDCEvt);
    if (rDCEvt.GetType() == DataChangedEventType::SETTINGS
        && (rDCEvt.GetFlags() & AllSettingsFlags::STYLE))
        PaintRail();
}

// Office's Backstage covers the ribbon too; the ribbon lives in the system
// window's border, outside our parent, so hide it while we are open.
void Backstage::ShowRibbon(bool bShow)
{
    if (SystemWindow* pSystemWindow = GetSystemWindow())
        if (const VclPtr<NotebookBar>& pNotebookBar = pSystemWindow->GetNotebookBar())
        {
            if (bShow)
                pNotebookBar->Show();
            else
                pNotebookBar->Hide();
            // the border window lays out the ribbon and the client area
            if (vcl::Window* pBorder = pSystemWindow->GetWindow(GetWindowType::Border))
                pBorder->Resize();
        }
}

void Backstage::Open(const OUString& rPage)
{
    Refresh();
    ShowRibbon(false);
    FitToParent();
    SetZOrder(nullptr, ZOrderFlags::First);
    Show();
    SelectPage(IsPage(rPage) ? rPage : u"home"_ustr);
    mxBack->grab_focus();
}

void Backstage::Close()
{
    if (!IsVisible())
        return;
    Hide();
    ShowRibbon(true);
    if (mxFrame.is())
        if (VclPtr<vcl::Window> pDoc = VCLUnoHelper::GetWindow(mxFrame->getComponentWindow()))
            pDoc->GrabFocus();
}

OUString Backstage::GetCurrentPage() const { return IsVisible() ? maPage : OUString(); }

void Backstage::SelectPage(const OUString& rPage)
{
    maPage = rPage;
    mxPages->set_current_page("tab_" + rPage); // vcl names pages by their tab label
    if (rPage == u"new" && !mbTemplatesFilled)
    {
        mbTemplatesFilled = true;
        FILTER_APPLICATION eApp = FILTER_APPLICATION::WRITER;
        if (maModule == u"com.sun.star.sheet.SpreadsheetDocument")
            eApp = FILTER_APPLICATION::CALC;
        else if (maModule == u"com.sun.star.presentation.PresentationDocument")
            eApp = FILTER_APPLICATION::IMPRESS;
        else if (maModule == u"com.sun.star.drawing.DrawingDocument")
            eApp = FILTER_APPLICATION::DRAW;
        mxTemplates->Populate();
        mxTemplates->filterItems(ViewFilter_Application(eApp));
        mxTemplates->showAllTemplates();
    }
    PaintRail();
}

void Backstage::Refresh()
{
    sal_Int32 nHour = tools::Time(tools::Time::SYSTEM).GetHour();
    mxGreeting->set_label(SfxResId(nHour < 12   ? STR_BACKSTAGE_GOOD_MORNING
                                   : nHour < 18 ? STR_BACKSTAGE_GOOD_AFTERNOON
                                                : STR_BACKSTAGE_GOOD_EVENING));
    mxHomeRecent->Reload();
    mxOpenRecent->Reload();
    mxPrinter->set_label(SfxResId(STR_BACKSTAGE_PRINTER).replaceFirst("%1", Printer::GetDefaultPrinterName()));
    RefreshInfo();
}

void Backstage::RefreshInfo()
{
    for (auto& [rKey, xLabel] : maProps)
        xLabel->set_label(OUString());
    if (!mxFrame.is() || !mxFrame->getController().is())
        return;
    uno::Reference<frame::XController> xController = mxFrame->getController();
    uno::Reference<frame::XModel> xModel = xController->getModel();

    OUString aName;
    if (uno::Reference<frame::XTitle> xTitle{ xModel, uno::UNO_QUERY })
        aName = xTitle->getTitle();
    mxInfoName->set_label(aName);

    OUString aURL = xModel.is() ? xModel->getURL() : OUString();
    if (aURL.isEmpty())
        mxInfoPath->set_label(SfxResId(STR_BACKSTAGE_NOT_SAVED));
    else
    {
        INetURLObject aObj(aURL);
        mxInfoPath->set_label(aObj.GetProtocol() == INetProtocol::File
                                  ? aObj.getFSysPath(FSysStyle::Detect)
                                  : aObj.GetMainURL(INetURLObject::DecodeMechanism::WithCharset));
        osl::DirectoryItem aItem;
        osl::FileStatus aStatus(osl_FileStatus_Mask_FileSize);
        if (aObj.GetProtocol() == INetProtocol::File
            && osl::DirectoryItem::get(aURL, aItem) == osl::FileBase::E_None
            && aItem.getFileStatus(aStatus) == osl::FileBase::E_None)
            maProps[u"size"_ustr]->set_label(FormatSize(aStatus.getFileSize()));
    }

    try
    {
        OUString aPagesLabel = SfxResId(STR_BACKSTAGE_PAGES);
        OUString aPages;
        if (uno::Reference<sheet::XSpreadsheetDocument> xCalc{ xModel, uno::UNO_QUERY })
        {
            aPagesLabel = SfxResId(STR_BACKSTAGE_SHEETS);
            aPages = OUString::number(xCalc->getSheets()->getElementNames().getLength());
        }
        else if (uno::Reference<drawing::XDrawPagesSupplier> xDraw{ xModel, uno::UNO_QUERY })
        {
            if (maModule == u"com.sun.star.presentation.PresentationDocument")
                aPagesLabel = SfxResId(STR_BACKSTAGE_SLIDES);
            aPages = OUString::number(xDraw->getDrawPages()->getCount());
        }
        else if (uno::Reference<beans::XPropertySet> xView{ xController, uno::UNO_QUERY })
        {
            if (xView->getPropertySetInfo()->hasPropertyByName(u"PageCount"_ustr))
                aPages = OUString::number(xView->getPropertyValue(u"PageCount"_ustr).get<sal_Int32>());
        }
        m_xBuilder->weld_label(u"prop_pages_label"_ustr)->set_label(aPagesLabel);
        maProps[u"pages"_ustr]->set_label(aPages);

        // The live count; DocumentStatistics only echoes what the file stored.
        bool bWords = false;
        if (uno::Reference<beans::XPropertySet> xProps{ xModel, uno::UNO_QUERY })
            if (xProps->getPropertySetInfo()->hasPropertyByName(u"WordCount"_ustr))
            {
                maProps[u"words"_ustr]->set_label(
                    OUString::number(xProps->getPropertyValue(u"WordCount"_ustr).get<sal_Int32>()));
                bWords = true;
            }
        m_xBuilder->weld_widget(u"prop_words_label"_ustr)->set_visible(bWords);
        maProps[u"words"_ustr]->set_visible(bWords);

        if (uno::Reference<document::XDocumentPropertiesSupplier> xSupplier{ xModel, uno::UNO_QUERY })
        {
            uno::Reference<document::XDocumentProperties> xDocProps = xSupplier->getDocumentProperties();
            maProps[u"title"_ustr]->set_label(xDocProps->getTitle());
            maProps[u"author"_ustr]->set_label(xDocProps->getAuthor());
            maProps[u"modified"_ustr]->set_label(FormatDate(xDocProps->getModificationDate()));
            maProps[u"created"_ustr]->set_label(FormatDate(xDocProps->getCreationDate()));
        }
    }
    catch (const uno::Exception&)
    {
    }
}

void Backstage::Dispatch(const OUString& rCommand)
{
    // Leave the Backstage first, as Office does; the command runs on the
    // document once this click is over.
    Close();
    auto* pPending = new PendingCommand{ mxFrame, rCommand };
    Application::PostUserEvent(LINK_NONMEMBER(nullptr, DispatchPending), pPending);
}

IMPL_LINK(Backstage, RailHdl, weld::Button&, rButton, void)
{
    OUString aId = rButton.get_buildable_name().copy(5); // "rail_"
    if (IsPage(aId))
    {
        if (aId == u"info" || aId == u"print")
            Refresh();
        SelectPage(aId);
    }
    else
        Dispatch(CommandFor(rButton.get_buildable_name()));
}

IMPL_LINK(Backstage, ActionHdl, weld::Button&, rButton, void)
{
    Dispatch(CommandFor(rButton.get_buildable_name()));
}

IMPL_LINK_NOARG(Backstage, BackHdl, weld::Button&, void) { Close(); }

IMPL_LINK_NOARG(Backstage, BlankHdl, weld::Button&, void) { Dispatch(u".uno:AddDirect"_ustr); }

IMPL_LINK_NOARG(Backstage, MoreHdl, weld::Button&, void) { SelectPage(u"new"_ustr); }

IMPL_LINK(Backstage, OpenTemplateHdl, const OUString&, rTemplatePath, void)
{
    Close();
    uno::Sequence<beans::PropertyValue> aArgs{
        comphelper::makePropertyValue(u"AsTemplate"_ustr, true),
        comphelper::makePropertyValue(u"MacroExecutionMode"_ustr,
                                      document::MacroExecMode::USE_CONFIG),
        comphelper::makePropertyValue(u"UpdateDocMode"_ustr,
                                      document::UpdateDocMode::ACCORDING_TO_CONFIG),
        comphelper::makePropertyValue(
            u"InteractionHandler"_ustr,
            task::InteractionHandler::createWithParent(comphelper::getProcessComponentContext(), nullptr))
    };
    try
    {
        uno::Reference<frame::XComponentLoader> xLoader(
            frame::Desktop::create(comphelper::getProcessComponentContext()), uno::UNO_QUERY_THROW);
        xLoader->loadComponentFromURL(rTemplatePath, u"_default"_ustr, 0, aArgs);
    }
    catch (const uno::Exception&)
    {
    }
}

bool Backstage::HandleEscape(const NotifyEvent& rNEvt)
{
    if (rNEvt.GetType() != NotifyEventType::KEYINPUT)
        return false;
    const vcl::KeyCode& rKey = rNEvt.GetKeyEvent()->GetKeyCode();
    if (rKey.GetCode() != KEY_ESCAPE || rKey.GetModifier())
        return false;
    Close();
    return true;
}

// Esc leaves the Backstage, from wherever in it the focus is.
bool Backstage::PreNotify(NotifyEvent& rNEvt)
{
    return HandleEscape(rNEvt) || InterimItemWindow::PreNotify(rNEvt);
}

bool Backstage::EventNotify(NotifyEvent& rNEvt)
{
    return HandleEscape(rNEvt) || InterimItemWindow::EventNotify(rNEvt);
}

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
