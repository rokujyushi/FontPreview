//----------------------------------------------------------------------------------
//	Font Preview Plugin (Grid + VF.object alias)
//----------------------------------------------------------------------------------
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <windowsx.h>
#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <algorithm>
#include <sstream>
#include <cmath>
#include <dwrite_3.h>
#include <d3d11.h>
#include <d2d1_1.h>
#include <d2d1_1helper.h>
#include <dxgi1_6.h>
#include <wrl/client.h>
#include <shlwapi.h>
#include <dwmapi.h>
using Microsoft::WRL::ComPtr;

#include "plugin2.h"
#include "logger2.h"
#include "AxisMapping.h"

#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dxgi.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwmapi.lib")

#define FontPreviewWindowName L"FontPreviewClient"
#define WM_DO_SET_FONT_OBJECT (WM_APP + 100)
#define IDC_FONT_GRID 1001
#define IDC_SEARCH_EDIT 1002
#define IDC_TYPE_FILTER 1003
#define IDC_BG_COLOR_BTN 1004
#define IDC_SAMPLE_TEXT_EDIT 1005
#define IDC_ADD_VF_BUTTON 1006
#define IDC_NAME_LABEL 1007
#define IDC_TYPE_LABEL 1008
#define IDC_AXIS_LABEL 1009
#define IDC_ADD_BUTTON 1010

constexpr int kGridCols = 2;
constexpr int kGridRows = 5;

enum class FontTypeFilter
{
	All = 0,
	System = 1,
	Folder = 2
};

EDIT_HANDLE *edit_handle = nullptr;
LOG_HANDLE *logger = nullptr;

ComPtr<IDWriteFactory7> g_dwriteFactory;
ComPtr<ID3D11Device> g_d3dDevice;
ComPtr<ID3D11DeviceContext> g_d3dContext;
ComPtr<IDXGISwapChain1> g_swapChain;
ComPtr<ID2D1Factory1> g_d2dFactory;
ComPtr<ID2D1Device> g_d2dDevice;
ComPtr<ID2D1DeviceContext> g_d2dContext;
ComPtr<ID2D1Bitmap1> g_d2dTarget;
D3D_FEATURE_LEVEL g_featureLevel = D3D_FEATURE_LEVEL_11_0;

struct FontItem
{
	std::wstring displayName;
	std::wstring filePath;
	bool isSystemFont = true;
	std::vector<std::string> axisTags;
	std::vector<std::pair<std::string, std::pair<float, float>>> axisRanges;
};

// -----------------------------------------------------------------
// Globals overview
// - Most UI HWNDs are stored in `g_hwnd*` globals so layout and
//   handlers can access controls across functions.
// - Graphics resources (D3D/D2D) are stored in g_d* globals.
// - `g_previewBgColor` governs only the GPU preview clear color.
// -----------------------------------------------------------------

std::vector<FontItem> g_fontList;
std::vector<int> g_filteredIndices;
int g_selectedFontIndex = -1;
FontTypeFilter g_filterType = FontTypeFilter::All;
std::wstring g_searchQuery;
HWND g_hwndGrid = nullptr;
HWND g_hwndPreview = nullptr;
HWND g_hwndSearch = nullptr;
HWND g_hwndType = nullptr;
HWND g_hwndBgBtn = nullptr;
HWND g_hwndSample = nullptr;
HWND g_hwndAddVF = nullptr;
HWND g_hwndAddText = nullptr;
HWND g_hwndNameLabel = nullptr;
HWND g_hwndTypeLabel = nullptr;
HWND g_hwndAxisLabel = nullptr;
std::wstring g_fontFolderPath;
COLORREF g_previewBgColor = RGB(255, 255, 255);
std::wstring g_sampleText = L"あいうABC123";
bool g_inRenderPreview = false;

static std::unordered_map<std::wstring, ComPtr<IDWriteFontCollection1>> g_externalFontCollections;

constexpr double kDefaultAliasSeconds = 1.1;
constexpr int kFallbackAliasFrames = 182;

static int ComputeAliasLengthFramesFromEditInfo(const EDIT_INFO &info, double seconds)
{
	if (!(seconds > 0.0))
		return 1;
	if (info.rate <= 0 || info.scale <= 0)
		return -1;
	double frames = seconds * (double)info.rate / (double)info.scale;
	int len = (int)std::ceil(frames);
	if (len < 1)
		len = 1;
	return len;
}

//---------------------------------------------------------------------
//	Utility helpers
//---------------------------------------------------------------------
// Convert a UTF-16 wide string to a UTF-8 encoded std::string.
// Used for interaction with APIs and alias text which expect UTF-8.
// Returns an empty std::string on failure or when input is empty.
std::string ToUtf8(const std::wstring &w)
{
	if (w.empty())
		return std::string();
	int len = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
	if (len <= 0)
		return std::string();
	std::string out(len - 1, '\0');
	WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, out.data(), len, nullptr, nullptr);
	return out;
}

std::wstring ToLower(const std::wstring &s)
{
	std::wstring r = s;
	std::transform(r.begin(), r.end(), r.begin(), ::towlower);
	return r;
}

static std::wstring ExtractFamilyName(const FontItem &item)
{
	std::wstring family = item.displayName;
	if (!item.isSystemFont)
	{
		size_t pos = family.find(L" [");
		if (pos != std::wstring::npos)
			family = family.substr(0, pos);
	}
	return family;
}

static HRESULT GetOrCreateExternalFontCollection(const std::wstring &filePath, IDWriteFontCollection1 **outCollection)
{
	if (!outCollection)
		return E_INVALIDARG;
	*outCollection = nullptr;
	if (!g_dwriteFactory)
		return E_FAIL;
	if (filePath.empty())
		return E_INVALIDARG;

	auto it = g_externalFontCollections.find(filePath);
	if (it != g_externalFontCollections.end() && it->second)
	{
		*outCollection = it->second.Get();
		(*outCollection)->AddRef();
		return S_OK;
	}

	ComPtr<IDWriteFontFile> fontFile;
	HRESULT hr = g_dwriteFactory->CreateFontFileReference(filePath.c_str(), nullptr, &fontFile);
	if (FAILED(hr) || !fontFile)
		return FAILED(hr) ? hr : E_FAIL;

	ComPtr<IDWriteFontSetBuilder> fontSetBuilder;
	hr = g_dwriteFactory->CreateFontSetBuilder(&fontSetBuilder);
	if (FAILED(hr) || !fontSetBuilder)
		return FAILED(hr) ? hr : E_FAIL;

	ComPtr<IDWriteFontSetBuilder1> fontSetBuilder1;
	hr = fontSetBuilder.As(&fontSetBuilder1);
	if (FAILED(hr) || !fontSetBuilder1)
		return FAILED(hr) ? hr : E_FAIL;

	fontSetBuilder1->AddFontFile(fontFile.Get());

	ComPtr<IDWriteFontSet> fontSet;
	hr = fontSetBuilder1->CreateFontSet(&fontSet);
	if (FAILED(hr) || !fontSet)
		return FAILED(hr) ? hr : E_FAIL;

	ComPtr<IDWriteFontCollection1> collection;
	hr = g_dwriteFactory->CreateFontCollectionFromFontSet(fontSet.Get(), &collection);
	if (FAILED(hr) || !collection)
		return FAILED(hr) ? hr : E_FAIL;

	g_externalFontCollections[filePath] = collection;
	*outCollection = collection.Detach();
	return S_OK;
}

//---------------------------------------------------------------------
//	Axis helpers
//---------------------------------------------------------------------
std::string TagToString(DWRITE_FONT_AXIS_TAG tag)
{
	char buf[5] = {0};
	buf[0] = (char)((tag >> 24) & 0xFF);
	buf[1] = (char)((tag >> 16) & 0xFF);
	buf[2] = (char)((tag >> 8) & 0xFF);
	buf[3] = (char)(tag & 0xFF);
	return std::string(buf);
}

	// Collect variation axis metadata (tags and ranges) from a DWrite font face.
	// Fills `item.axisTags` and `item.axisRanges` when the font supports variations.
	void CollectFontAxes(FontItem &item, IDWriteFontFace5 *fontFace)
{
	if (!fontFace)
		return;
	item.axisTags.clear();
	item.axisRanges.clear();

	if (!fontFace->HasVariations())
		return;

	ComPtr<IDWriteFontResource> fontResource;
	HRESULT hr = fontFace->GetFontResource(&fontResource);
	if (FAILED(hr) || !fontResource)
		return;

	UINT32 axisCount = fontResource->GetFontAxisCount();
	if (axisCount == 0)
		return;

	std::vector<DWRITE_FONT_AXIS_VALUE> defaultAxisValues(axisCount);
	std::vector<DWRITE_FONT_AXIS_RANGE> axisRanges(axisCount);

	hr = fontResource->GetDefaultFontAxisValues(defaultAxisValues.data(), axisCount);
	if (FAILED(hr))
		return;
	hr = fontResource->GetFontAxisRanges(axisRanges.data(), axisCount);
	if (FAILED(hr))
		return;

	for (UINT32 i = 0; i < axisCount; i++)
	{
		std::string tagStr = TagToString(defaultAxisValues[i].axisTag);
		item.axisTags.push_back(tagStr);
		item.axisRanges.push_back({tagStr, {axisRanges[i].minValue, axisRanges[i].maxValue}});
	}
}

std::wstring BuildAxisTooltip(const FontItem &item)
{
	if (item.axisTags.empty())
	{
		return L"可変フォント軸はありません";
	}
	std::wstring tip;
	int lines = 0;
	const int kMaxLines = 15;
	for (const auto &axis : item.axisRanges)
	{
		if (lines >= kMaxLines)
		{
			tip += L"...";
			break;
		}
		const std::string &tag = axis.first;
		float minV = axis.second.first;
		float maxV = axis.second.second;
		std::string human = AxisMapping::GetAxisHumanName(tag);
		std::wstring tagW(tag.begin(), tag.end());
		std::wstring humanW(human.begin(), human.end());
		wchar_t buf[128];
		swprintf_s(buf, L"%ls %ls %.1f-%.1f", tagW.c_str(), humanW.c_str(), minV, maxV);
		if (!tip.empty())
			tip += L"\n";
		tip += buf;
		lines++;
	}
	return tip;
}

std::wstring BuildAxisTagLine(const FontItem &item)
{
	if (item.axisTags.empty())
		return L"";
	std::wstring line;
	const size_t kMaxTags = 15;
	size_t count = 0;
	for (const auto &tag : item.axisTags)
	{
		if (count >= kMaxTags)
		{
			line += L" ...";
			break;
		}
		std::wstring tagW(tag.begin(), tag.end());
		if (!line.empty())
			line += L" ";
		line += tagW;
		count++;
	}
	return line;
}

//---------------------------------------------------------------------
//	UI helpers
//---------------------------------------------------------------------
void SyncSampleTextFromEdit()
{
	if (!g_hwndSample)
		return;
	wchar_t buf[512] = {0};
	int len = GetWindowTextW(g_hwndSample, buf, _countof(buf));
	if (len >= 0)
	{
		g_sampleText.assign(buf, len);
		if (logger)
		{
			wchar_t logbuf[160];
			swprintf_s(logbuf, L"SyncSampleTextFromEdit: len=%d", len);
			logger->verbose(logger, logbuf);
		}
	}
}

//---------------------------------------------------------------------
//	Graphics init
//---------------------------------------------------------------------
HRESULT InitializeGraphics()
{
	HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory7), (IUnknown **)&g_dwriteFactory);
	if (FAILED(hr))
	{
		if (logger)
		{
			wchar_t msg[128];
			swprintf_s(msg, L"DWrite factory failed: 0x%08x", hr);
			logger->warn(logger, msg);
		}
		return hr;
	}
	if (logger)
		logger->log(logger, L"InitializeGraphics: DWrite initialized");
	return S_OK;
}

//---------------------------------------------------------------------
//	Font enumeration helpers
//---------------------------------------------------------------------
std::wstring GetDefaultFontFolder()
{
	wchar_t modulePath[MAX_PATH] = {0};
	HMODULE hMod = NULL;
	if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCWSTR)&GetDefaultFontFolder, &hMod))
	{
		if (GetModuleFileNameW(hMod, modulePath, MAX_PATH) == 0)
		{
			GetModuleFileNameW(NULL, modulePath, MAX_PATH);
		}
		FreeLibrary(hMod);
	}
	else
	{
		GetModuleFileNameW(NULL, modulePath, MAX_PATH);
	}
	PathRemoveFileSpec(modulePath);
	std::wstring fontFolder = modulePath;
	fontFolder += L"\\Fonts";
	return fontFolder;
}

void EnumerateFolderFonts(const std::wstring &folderPath, std::unordered_set<std::wstring> &seenNames)
{
	if (!g_dwriteFactory)
		return;

	std::wstring searchPath = folderPath + L"\\*.*";
	WIN32_FIND_DATAW findData{};
	HANDLE hFind = FindFirstFileW(searchPath.c_str(), &findData);
	if (hFind == INVALID_HANDLE_VALUE)
		return;

	do
	{
		if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			continue;
		std::wstring fileName = findData.cFileName;
		std::wstring lowerName = ToLower(fileName);
		bool isFontFile = false;
		if (lowerName.length() >= 4)
		{
			std::wstring ext = lowerName.substr(lowerName.length() - 4);
			isFontFile = (ext == L".ttf" || ext == L".otf" || ext == L".ttc");
		}
		if (!isFontFile)
			continue;

		ComPtr<IDWriteFontFile> fontFile;
		HRESULT hr = g_dwriteFactory->CreateFontFileReference((folderPath + L"\\" + fileName).c_str(), nullptr, &fontFile);
		if (FAILED(hr))
			continue;

		BOOL isSupported = FALSE;
		DWRITE_FONT_FILE_TYPE fontFileType;
		DWRITE_FONT_FACE_TYPE fontFaceType;
		UINT32 numberOfFaces = 0;
		hr = fontFile->Analyze(&isSupported, &fontFileType, &fontFaceType, &numberOfFaces);
		if (FAILED(hr) || !isSupported)
			continue;

		ComPtr<IDWriteFontSetBuilder> fontSetBuilder;
		if (FAILED(g_dwriteFactory->CreateFontSetBuilder(&fontSetBuilder)))
			continue;
		ComPtr<IDWriteFontSetBuilder1> fontSetBuilder1;
		if (FAILED(fontSetBuilder.As(&fontSetBuilder1)) || !fontSetBuilder1)
			continue;
		fontSetBuilder1->AddFontFile(fontFile.Get());

		ComPtr<IDWriteFontSet> fontSet;
		if (FAILED(fontSetBuilder1->CreateFontSet(&fontSet)))
			continue;
		ComPtr<IDWriteFontSet1> fontSet1;
		fontSet.As(&fontSet1);
		UINT32 fontCount = fontSet ? fontSet->GetFontCount() : 0;

		for (UINT32 i = 0; i < fontCount; i++)
		{
			ComPtr<IDWriteLocalizedStrings> familyNames;
			BOOL exists = FALSE;
			if (!fontSet1)
				continue;
			if (FAILED(fontSet1->GetPropertyValues(i, DWRITE_FONT_PROPERTY_ID_FAMILY_NAME, &exists, &familyNames)) || !exists)
				continue;
			UINT32 index = 0;
			BOOL localeExists = false;
			if (FAILED(familyNames->FindLocaleName(L"ja-jp", &index, &localeExists)) || !localeExists)
			{
				familyNames->FindLocaleName(L"en-us", &index, &localeExists);
			}
			if (!localeExists)
				index = 0;
			UINT32 length = 0;
			if (FAILED(familyNames->GetStringLength(index, &length)))
				continue;
			std::wstring familyName(length + 1, L'\0');
			if (FAILED(familyNames->GetString(index, &familyName[0], length + 1)))
				continue;
			familyName.resize(length);

			FontItem item;
			item.displayName = familyName + L" [" + fileName + L"]";
			item.filePath = folderPath + L"\\" + fileName;
			item.isSystemFont = false;

			ComPtr<IDWriteFontFace> tempFace;
			ComPtr<IDWriteFontFace5> face5;
			IDWriteFontFile *files[] = {fontFile.Get()};
			if (SUCCEEDED(g_dwriteFactory->CreateFontFace(fontFaceType, 1, files, 0, DWRITE_FONT_SIMULATIONS_NONE, &tempFace)))
			{
				tempFace.As(&face5);
			}
			if (face5)
				CollectFontAxes(item, face5.Get());

			std::wstring key = ToLower(item.displayName);
			if (seenNames.find(key) != seenNames.end())
			{
				if (logger)
				{
					std::wstring msg = L"Font collision (keep first): " + item.displayName;
					logger->log(logger, msg.c_str());
				}
				continue;
			}
			seenNames.insert(key);
			g_fontList.push_back(item);
		}
	} while (FindNextFileW(hFind, &findData));
	FindClose(hFind);
}

void EnumerateFonts()
{
	g_fontList.clear();
	std::unordered_set<std::wstring> seenNames;
	if (!g_dwriteFactory)
		return;
	if (logger)
		logger->info(logger, L"EnumerateFonts: start");

	ComPtr<IDWriteFontCollection> fontCollection;
	HRESULT hr = g_dwriteFactory->GetSystemFontCollection(&fontCollection);
	if (FAILED(hr))
		return;

	UINT32 familyCount = fontCollection->GetFontFamilyCount();
	for (UINT32 i = 0; i < familyCount; i++)
	{
		ComPtr<IDWriteFontFamily> fontFamily;
		if (FAILED(fontCollection->GetFontFamily(i, &fontFamily)))
			continue;
		ComPtr<IDWriteLocalizedStrings> familyNames;
		if (FAILED(fontFamily->GetFamilyNames(&familyNames)))
			continue;
		UINT32 index = 0;
		BOOL exists = false;
		if (FAILED(familyNames->FindLocaleName(L"ja-jp", &index, &exists)) || !exists)
		{
			familyNames->FindLocaleName(L"en-us", &index, &exists);
		}
		if (!exists)
			index = 0;
		UINT32 length = 0;
		if (FAILED(familyNames->GetStringLength(index, &length)))
			continue;
		std::wstring familyName(length + 1, L'\0');
		if (FAILED(familyNames->GetString(index, &familyName[0], length + 1)))
			continue;
		familyName.resize(length);

		FontItem item;
		item.displayName = familyName;
		item.filePath.clear();
		item.isSystemFont = true;

		ComPtr<IDWriteFont> matchFont;
		if (SUCCEEDED(fontFamily->GetFirstMatchingFont(DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STRETCH_NORMAL, DWRITE_FONT_STYLE_NORMAL, &matchFont)))
		{
			ComPtr<IDWriteFontFace> baseFace;
			ComPtr<IDWriteFontFace5> face5;
			if (SUCCEEDED(matchFont->CreateFontFace(&baseFace)))
				baseFace.As(&face5);
			if (face5)
				CollectFontAxes(item, face5.Get());
		}

		std::wstring key = ToLower(item.displayName);
		if (seenNames.insert(key).second)
		{
			g_fontList.push_back(item);
			if (logger)
			{
				std::wstring msg = L"EnumerateFonts: added " + item.displayName;
				logger->verbose(logger, msg.c_str());
			}
		}
	}

	if (g_fontFolderPath.empty())
		g_fontFolderPath = GetDefaultFontFolder();
	if (PathFileExistsW(g_fontFolderPath.c_str()))
	{
		EnumerateFolderFonts(g_fontFolderPath, seenNames);
	}
	if (logger)
	{
		wchar_t buf[128];
		swprintf_s(buf, L"EnumerateFonts: total fonts=%d", (int)g_fontList.size());
		logger->info(logger, buf);
	}
}

//---------------------------------------------------------------------
//	Filtering and selection
//---------------------------------------------------------------------
void UpdateDetailPanel();
void RedrawGrid();
void RebuildListViewItems();
void RenderPreview(const wchar_t *reason = L"");
bool EnsurePreviewDevice();
bool CreateOrResizeSwapChain(HWND hwnd, int width, int height);
void ReleasePreviewTarget();

void ApplyFilter()
{
	g_filteredIndices.clear();
	std::wstring qLower = ToLower(g_searchQuery);
	if (logger)
	{
		std::wstring msg = L"ApplyFilter: query='" + g_searchQuery + L"'";
		logger->verbose(logger, msg.c_str());
	}
	for (size_t i = 0; i < g_fontList.size(); i++)
	{
		const auto &item = g_fontList[i];
		if (g_filterType == FontTypeFilter::System && !item.isSystemFont)
			continue;
		if (g_filterType == FontTypeFilter::Folder && item.isSystemFont)
			continue;
		if (!qLower.empty())
		{
			std::wstring nameLower = ToLower(item.displayName);
			if (nameLower.find(qLower) == std::wstring::npos)
				continue;
		}
		g_filteredIndices.push_back((int)i);
	}
	RebuildListViewItems();
	if (!g_filteredIndices.empty())
	{
		if (std::find(g_filteredIndices.begin(), g_filteredIndices.end(), g_selectedFontIndex) == g_filteredIndices.end())
		{
			g_selectedFontIndex = g_filteredIndices.front();
		}
	}
	else
	{
		g_selectedFontIndex = -1;
	}
	UpdateDetailPanel();
	RedrawGrid();
	RenderPreview(L"ApplyFilter");
	if (logger)
	{
		wchar_t buf[128];
		swprintf_s(buf, L"ApplyFilter: filtered=%d", (int)g_filteredIndices.size());
		logger->info(logger, buf);
	}
}

void RebuildListViewItems()
{
	if (!g_hwndGrid)
		return;
	ListView_DeleteAllItems(g_hwndGrid);
	LVITEMW item{};
	item.mask = LVIF_TEXT | LVIF_PARAM;
	for (size_t i = 0; i < g_filteredIndices.size(); i++)
	{
		item.iItem = (int)i;
		item.lParam = g_filteredIndices[i];
		int fontIdx = g_filteredIndices[i];
		if (fontIdx >= 0 && fontIdx < (int)g_fontList.size())
		{
			item.pszText = const_cast<wchar_t *>(g_fontList[fontIdx].displayName.c_str());
		}
		else
		{
			item.pszText = const_cast<wchar_t *>(L"");
		}
		ListView_InsertItem(g_hwndGrid, &item);
	}
	if (g_selectedFontIndex >= 0)
	{
		for (size_t i = 0; i < g_filteredIndices.size(); i++)
		{
			if (g_filteredIndices[i] == g_selectedFontIndex)
			{
				ListView_SetItemState(g_hwndGrid, (int)i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
				ListView_EnsureVisible(g_hwndGrid, (int)i, FALSE);
				break;
			}
		}
	}
	if (logger)
	{
		wchar_t buf[128];
		swprintf_s(buf, L"RebuildListViewItems: items=%d", (int)g_filteredIndices.size());
		logger->verbose(logger, buf);
	}
}

void RedrawGrid()
{
	if (g_hwndGrid)
	{
		RedrawWindow(g_hwndGrid, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_UPDATENOW);
		if (logger)
			logger->verbose(logger, L"RedrawGrid: invalidated ListView");
	}
}

bool EnsurePreviewDevice()
{
	if (g_d3dDevice && g_d3dContext && g_d2dFactory && g_d2dDevice && g_d2dContext)
		return true;
	UINT flags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
#if defined(_DEBUG)
	flags |= D3D11_CREATE_DEVICE_DEBUG;
#endif
	D3D_FEATURE_LEVEL levels[] = {
		D3D_FEATURE_LEVEL_11_1,
		D3D_FEATURE_LEVEL_11_0,
		D3D_FEATURE_LEVEL_10_1,
		D3D_FEATURE_LEVEL_10_0,
	};
	ComPtr<ID3D11Device> device;
	ComPtr<ID3D11DeviceContext> context;
	HRESULT hr = D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, flags, levels, _countof(levels), D3D11_SDK_VERSION, &device, &g_featureLevel, &context);
	if (FAILED(hr))
	{
		if (logger)
		{
			wchar_t buf[128];
			swprintf_s(buf, L"D3D11CreateDevice failed: 0x%08x", hr);
			logger->warn(logger, buf);
		}
		return false;
	}
	g_d3dDevice = device;
	g_d3dContext = context;
	ComPtr<IDXGIDevice> dxgiDevice;
	if (FAILED(g_d3dDevice.As(&dxgiDevice)))
		return false;
	D2D1_FACTORY_OPTIONS opts{};
	if (FAILED(D2D1CreateFactory(D2D1_FACTORY_TYPE_MULTI_THREADED, __uuidof(ID2D1Factory1), &opts, &g_d2dFactory)))
		return false;
	if (FAILED(g_d2dFactory->CreateDevice(dxgiDevice.Get(), &g_d2dDevice)))
		return false;
	if (FAILED(g_d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, &g_d2dContext)))
		return false;
	if (logger)
		logger->info(logger, L"Preview: D3D11/D2D initialized");
	return true;
}

void ReleasePreviewTarget()
{
	if (g_d2dContext)
	{
		g_d2dContext->SetTarget(nullptr);
	}
	g_d2dTarget.Reset();
}

bool CreateOrResizeSwapChain(HWND hwnd, int width, int height)
{
	if (!hwnd || width <= 0 || height <= 0)
		return false;
	if (logger)
	{
		wchar_t buf[160];
		swprintf_s(buf, L"CreateOrResizeSwapChain: hwnd=%p size=%dx%d", hwnd, width, height);
		logger->verbose(logger, buf);
	}
	if (!EnsurePreviewDevice())
	{
		if (logger)
			logger->warn(logger, L"CreateOrResizeSwapChain: EnsurePreviewDevice failed");
		return false;
	}
	ReleasePreviewTarget();
	ComPtr<IDXGIDevice> dxgiDevice;
	if (FAILED(g_d3dDevice.As(&dxgiDevice)))
	{
		if (logger)
			logger->warn(logger, L"CreateOrResizeSwapChain: As IDXGIDevice failed");
		return false;
	}
	ComPtr<IDXGIAdapter> adapter;
	if (FAILED(dxgiDevice->GetAdapter(&adapter)))
	{
		if (logger)
			logger->warn(logger, L"CreateOrResizeSwapChain: GetAdapter failed");
		return false;
	}
	ComPtr<IDXGIFactory2> factory;
	if (FAILED(adapter->GetParent(__uuidof(IDXGIFactory2), &factory)))
	{
		if (logger)
			logger->warn(logger, L"CreateOrResizeSwapChain: GetParent IDXGIFactory2 failed");
		return false;
	}

	if (!g_swapChain)
	{
		DXGI_SWAP_CHAIN_DESC1 desc{};
		desc.Format = DXGI_FORMAT_B8G8R8A8_UNORM;
		desc.Width = width;
		desc.Height = height;
		desc.BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT;
		desc.SampleDesc.Count = 1;
		desc.BufferCount = 2;
		desc.SwapEffect = DXGI_SWAP_EFFECT_FLIP_DISCARD;
		desc.AlphaMode = DXGI_ALPHA_MODE_IGNORE;
		HRESULT hr = factory->CreateSwapChainForHwnd(g_d3dDevice.Get(), hwnd, &desc, nullptr, nullptr, &g_swapChain);
		if (FAILED(hr))
		{
			if (logger)
			{
				wchar_t buf[160];
				swprintf_s(buf, L"CreateSwapChainForHwnd failed 0x%08x", hr);
				logger->warn(logger, buf);
			}
			return false;
		}
		factory->MakeWindowAssociation(hwnd, DXGI_MWA_NO_ALT_ENTER);
	}
	else
	{
		HRESULT hr = g_swapChain->ResizeBuffers(0, width, height, DXGI_FORMAT_UNKNOWN, 0);
		if (FAILED(hr))
		{
			if (logger)
			{
				wchar_t buf[160];
				swprintf_s(buf, L"ResizeBuffers failed 0x%08x", hr);
				logger->warn(logger, buf);
			}
			// Drop the swapchain to force recreate next time; stale target may hold references
			g_swapChain.Reset();
			return false;
		}
	}

	ComPtr<IDXGISurface> surface;
	HRESULT hr = g_swapChain->GetBuffer(0, __uuidof(IDXGISurface), &surface);
	if (FAILED(hr))
	{
		if (logger)
		{
			wchar_t buf[160];
			swprintf_s(buf, L"GetBuffer failed 0x%08x", hr);
			logger->warn(logger, buf);
		}
		return false;
	}
	D2D1_BITMAP_PROPERTIES1 props = D2D1::BitmapProperties1(
		D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW,
		D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE),
		96.0f, 96.0f);
	hr = g_d2dContext->CreateBitmapFromDxgiSurface(surface.Get(), &props, &g_d2dTarget);
	if (FAILED(hr))
	{
		if (logger)
		{
			wchar_t buf[160];
			swprintf_s(buf, L"CreateBitmapFromDxgiSurface failed 0x%08x", hr);
			logger->warn(logger, buf);
		}
		return false;
	}
	g_d2dContext->SetTarget(g_d2dTarget.Get());
	if (logger)
		logger->verbose(logger, L"CreateOrResizeSwapChain: target set successfully");
	return true;
}

// Render the GPU-backed preview area into the preview HWND's swapchain.
// This function creates/resizes the swap chain as needed, clears with
// `g_previewBgColor`, draws sample text using DirectWrite/Direct2D,
// and presents the buffer. It guards against re-entrant calls.
void RenderPreview(const wchar_t *reason)
{
	if (g_inRenderPreview)
		return;
	g_inRenderPreview = true;
	struct Guard
	{
		bool &flag;
		~Guard() { flag = false; }
	} guard{g_inRenderPreview};
	if (!g_hwndPreview)
	{
		if (logger)
			logger->warn(logger, L"RenderPreview: no preview hwnd");
		return;
	}
	SyncSampleTextFromEdit();
	RECT rc{};
	GetClientRect(g_hwndPreview, &rc);
	int w = rc.right - rc.left;
	int h = rc.bottom - rc.top;
	if (logger)
	{
		wchar_t buf[160];
		swprintf_s(buf, L"RenderPreview start[%ls]: size=%dx%d selected=%d", reason ? reason : L"", w, h, g_selectedFontIndex);
		logger->verbose(logger, buf);
	}
	if (w <= 0 || h <= 0)
		return;
	if (!CreateOrResizeSwapChain(g_hwndPreview, w, h))
	{
		if (logger)
			logger->warn(logger, L"RenderPreview: swapchain create/resize failed");
		return;
	}
	if (!g_d2dContext || !g_d2dTarget)
	{
		if (logger)
			logger->warn(logger, L"RenderPreview: D2D context/target missing");
		return;
	}

	float bgR = GetRValue(g_previewBgColor) / 255.0f;
	float bgG = GetGValue(g_previewBgColor) / 255.0f;
	float bgB = GetBValue(g_previewBgColor) / 255.0f;
	if (logger)
	{
		wchar_t buf[160];
		swprintf_s(buf, L"RenderPreview bg color=%06x", g_previewBgColor & 0xFFFFFF);
		logger->verbose(logger, buf);
	}
	g_d2dContext->BeginDraw();
	g_d2dContext->Clear(D2D1::ColorF(bgR, bgG, bgB, 1.0f));

	std::wstring sample = g_sampleText;
	if (sample.empty())
	{
		sample = L"あいうABC123";
		if (logger)
			logger->warn(logger, L"RenderPreview: sample text empty, using fallback");
	}

	int fontIdx = g_selectedFontIndex;
	if (fontIdx >= 0 && fontIdx < (int)g_fontList.size())
	{
		const auto &item = g_fontList[fontIdx];
		std::wstring family = ExtractFamilyName(item);
		ComPtr<IDWriteFontCollection1> externalCollection;
		IDWriteFontCollection *collectionForCreate = nullptr;
		if (!item.isSystemFont)
		{
			HRESULT colHr = GetOrCreateExternalFontCollection(item.filePath, &externalCollection);
			if (SUCCEEDED(colHr) && externalCollection)
				collectionForCreate = externalCollection.Get();
			else if (logger)
				logger->warn(logger, L"RenderPreview: external font collection create failed; falling back to system collection");
		}
		ComPtr<IDWriteTextFormat> format;
		HRESULT hrPrimary = g_dwriteFactory->CreateTextFormat(family.c_str(), collectionForCreate, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 48.0f, L"ja-jp", &format);
		HRESULT hrFallback = hrPrimary;
		if (FAILED(hrPrimary))
		{
			hrFallback = g_dwriteFactory->CreateTextFormat(L"Segoe UI", nullptr, DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 48.0f, L"ja-jp", &format);
			if (logger)
			{
				wchar_t buf[200];
				swprintf_s(buf, L"RenderPreview: primary format failed 0x%08x, fallback hr=0x%08x", hrPrimary, hrFallback);
				logger->warn(logger, buf);
			}
		}
		HRESULT hr = FAILED(hrPrimary) ? hrFallback : hrPrimary;
		if (SUCCEEDED(hr))
		{
			ComPtr<ID2D1SolidColorBrush> textBrush;
			if (SUCCEEDED(g_d2dContext->CreateSolidColorBrush(D2D1::ColorF(0, 0, 0, 1), &textBrush)))
			{
				D2D1_RECT_F layout = D2D1::RectF(10.0f, 10.0f, (FLOAT)w - 10.0f, (FLOAT)h - 10.0f);
				g_d2dContext->DrawTextW(sample.c_str(), (UINT32)sample.size(), format.Get(), layout, textBrush.Get(), D2D1_DRAW_TEXT_OPTIONS_NONE, DWRITE_MEASURING_MODE_NATURAL);
				if (logger)
				{
					wchar_t buf[200];
					swprintf_s(buf, L"RenderPreview: drew text len=%u family=%ls", (unsigned)sample.size(), family.c_str());
					logger->verbose(logger, buf);
				}
			}
		}
	}
	else
	{
		if (logger)
			logger->warn(logger, L"RenderPreview: no valid font index to draw");
	}

	HRESULT endHr = g_d2dContext->EndDraw();
	if (FAILED(endHr) && logger)
	{
		wchar_t buf[128];
		swprintf_s(buf, L"RenderPreview: EndDraw failed 0x%08x", endHr);
		logger->warn(logger, buf);
	}
	DXGI_PRESENT_PARAMETERS params{};
	HRESULT presentHr = g_swapChain->Present1(1, 0, &params);
	if (FAILED(presentHr) && logger)
	{
		wchar_t buf[128];
		swprintf_s(buf, L"RenderPreview: Present1 failed 0x%08x", presentHr);
		logger->warn(logger, buf);
	}
}

void UpdateDetailPanel()
{
	if (!g_hwndNameLabel || !g_hwndTypeLabel || !g_hwndAxisLabel)
		return;
	if (g_selectedFontIndex < 0 || g_selectedFontIndex >= (int)g_fontList.size())
	{
		SetWindowTextW(g_hwndNameLabel, L"フォント未選択");
		SetWindowTextW(g_hwndTypeLabel, L"");
		SetWindowTextW(g_hwndAxisLabel, L"");
		return;
	}
	const auto &item = g_fontList[g_selectedFontIndex];
	SetWindowTextW(g_hwndNameLabel, item.displayName.c_str());
	SetWindowTextW(g_hwndTypeLabel, item.isSystemFont ? L"システムフォント" : L"外部フォント");
	std::wstring axis = BuildAxisTooltip(item);
	SetWindowTextW(g_hwndAxisLabel, axis.c_str());
}

//---------------------------------------------------------------------
//	Alias generation (VF.object baseline)
//---------------------------------------------------------------------
std::string BuildVFAliasFromSelection(const FontItem &item, const std::wstring &text, int frameLength)
{
	if (frameLength <= 0)
		frameLength = kFallbackAliasFrames;
	std::wstring fontValue = item.isSystemFont ? item.displayName : item.filePath;
	std::ostringstream alias;
	alias << "[Object]\n";
	alias << "frame=0," << frameLength << "\n";
	alias << "[Object.0]\n";
	alias << "effect.name=Variable Font Text\n";
	alias << "フォントファイル=" << (item.isSystemFont ? "" : ToUtf8(fontValue)) << "\n";
	alias << "フォント=" << (item.isSystemFont ? ToUtf8(fontValue) : "") << "\n";
	alias << "サイズ=80.0\n";
	alias << "文字色=ffffff\n";
	alias << "B=0\n";
	alias << "I=0\n";
	alias << "字間=0.0\n";
	alias << "影設定.hide=1\n";
	alias << "影を表示=0\n";
	alias << "影色=000000\n";
	alias << "影X=0.0\n";
	alias << "影Y=0.0\n";
	alias << "影濃度=100\n";
	alias << "影ぼかし=0.0\n";
	alias << "縁取り設定.hide=1\n";
	alias << "縁取りを表示=0\n";
	alias << "縁取り色=000000\n";
	alias << "縁取り幅=5.0\n";
	alias << "縁取りスタイル=丸\n";
	alias << "切り抜き=0\n";
	alias << "Weight=400\n";
	alias << "Width=100\n";
	alias << "Slant=0.0\n";
	alias << "Optical Size=12.0\n";
	alias << "Italic Axis=0.0\n";
	alias << "Grade (GRAD)=0.0\n";
	alias << "XTRA=0\n";
	alias << "XOPQ=0\n";
	alias << "YOPQ=0\n";
	alias << "YTLC=0\n";
	alias << "YTUC=0\n";
	alias << "YTAS=0\n";
	alias << "YTDE=0\n";
	alias << "YTFI=0\n";
	alias << "軸更新モード=リアルタイム\n";
	alias << "横幅=0\n";
	alias << "縦幅=0\n";
	alias << "文字揃え=中央揃え[中]\n";
	alias << "行間=0.0\n";
	alias << "アニメーション.hide=1\n";
	alias << "表示速度=0.0\n";
	alias << "文字毎に個別オブジェクト=1\n";
	alias << "テキスト=" << ToUtf8(text) << "\n";
	alias << "[Object.1]\n";
	alias << "effect.name=標準描画\n";
	alias << "X=0.00\n";
	alias << "Y=0.00\n";
	alias << "Z=0.00\n";
	alias << "Group=1\n";
	alias << "中心X=0.00\n";
	alias << "中心Y=0.00\n";
	alias << "中心Z=0.00\n";
	alias << "X軸回転=0.00\n";
	alias << "Y軸回転=0.00\n";
	alias << "Z軸回転=0.00\n";
	alias << "拡大率=100.000\n";
	alias << "縦横比=0.000\n";
	alias << "透明度=0.00\n";
	alias << "合成モード=通常\n";
	return alias.str();
}
//---------------------------------------------------------------------
//	Alias generation (Text.object baseline)
//---------------------------------------------------------------------
std::string BuildAliasFromSelection(const FontItem &item, const std::wstring &text, int frameLength)
{
	if (frameLength <= 0)
		frameLength = kFallbackAliasFrames;
	std::ostringstream alias;
	alias << "[Object]\n";
	alias << "frame=0," << frameLength << "\n";
	alias << "[Object.0]\n";
	alias << "effect.name=テキスト\n";
	alias << "サイズ=80.0\n";
	alias << "字間=0.00\n";
	alias << "行間=0.00\n";
	alias << "表示速度=0.00\n";
	alias << "フォント=" << ToUtf8(item.displayName) << "\n";
	alias << "文字色=ffffff\n";
	alias << "影・縁色=000000\n";
	alias << "文字装飾=標準文字\n";
	alias << "文字揃え=中央揃え[中]\n";
	alias << "B=0\n";
	alias << "I=0\n";
	alias << "テキスト=" << ToUtf8(text) << "\n";
	alias << "文字毎に個別オブジェクト=0\n";
	alias << "自動スクロール=0\n";
	alias << "移動座標上に表示=0\n";
	alias << "オブジェクトの長さを自動調節=0\n";
	alias << "[Object.1]\n";
	alias << "effect.name=標準描画\n";
	alias << "X=0.00\n";
	alias << "Y=0.00\n";
	alias << "Z=0.00\n";
	alias << "Group=1\n";
	alias << "中心X=0.00\n";
	alias << "中心Y=0.00\n";
	alias << "中心Z=0.00\n";
	alias << "X軸回転=0.00\n";
	alias << "Y軸回転=0.00\n";
	alias << "Z軸回転=0.00\n";
	alias << "拡大率=100.000\n";
	alias << "縦横比=0.000\n";
	alias << "透明度=0.00\n";
	alias << "合成モード=通常\n";
	return alias.str();
}

bool CreateVariableFontObject(UINT flags)
{
	if (g_selectedFontIndex < 0 || g_selectedFontIndex >= (int)g_fontList.size())
	{
		if (logger)
			logger->warn(logger, L"フォントが選択されていません");
		return false;
	}
	if (!edit_handle)
	{
		if (logger)
			logger->error(logger, L"編集ハンドルが利用できません");
		return false;
	}
	const auto &item = g_fontList[g_selectedFontIndex];

	int frameLength = kFallbackAliasFrames;
	if (edit_handle->get_edit_info)
	{
		EDIT_INFO info{};
		edit_handle->get_edit_info(&info, sizeof(info));
		int computed = ComputeAliasLengthFramesFromEditInfo(info, kDefaultAliasSeconds);
		if (computed > 0)
			frameLength = computed;
		if (logger)
		{
			wchar_t buf[192];
			swprintf_s(buf, L"Alias length: seconds=%.3f rate=%d scale=%d => frames=%d", kDefaultAliasSeconds, info.rate, info.scale, frameLength);
			logger->verbose(logger, buf);
		}
	}

	std::string alias;
	switch (flags)
	{
	case IDC_ADD_VF_BUTTON:
		alias = BuildVFAliasFromSelection(item, g_sampleText, frameLength);
		break;
	case IDC_ADD_BUTTON:
		alias = BuildAliasFromSelection(item, g_sampleText, frameLength);
		break;
	default:
		break;
	}

	struct CreateAliasParam
	{
		std::string aliasText;
		bool created = false;
	};
	CreateAliasParam param{alias, false};

	bool called = edit_handle->call_edit_section_param(&param, [](void *p, EDIT_SECTION *edit)
													   {
		auto* ctx = static_cast<CreateAliasParam*>(p);
		if (!ctx || !edit || !edit->create_object_from_alias) return;
		int layer = edit->info ? edit->info->layer : 0;
		int frame = edit->info ? edit->info->frame : 0;
		ctx->created = edit->create_object_from_alias(ctx->aliasText.c_str(), layer, frame, 0) != nullptr; });

	bool ok = called && param.created;
	if (!ok)
	{
		if (logger)
			logger->error(logger, L"オブジェクト作成に失敗しました");
		if (logger)
			logger->warn(logger, L"create_object_from_alias failed");
	}
	else if (logger)
	{
		logger->log(logger, L"Variable Font Text object created from FontPreview");
	}
	return ok;
}

// Update the selected object(s) in the host editor with the currently
// selected font from the preview UI. This invokes `call_edit_section_param`
// to run on the host's edit section thread and set effect item values.
// Returns true if the host update succeeded.
bool SetFontTextObject()
{
	if (g_selectedFontIndex < 0 || g_selectedFontIndex >= (int)g_fontList.size())
		return false;
	const auto &item = g_fontList[g_selectedFontIndex];

	if (logger)
	{
		wchar_t buf[256];
		swprintf_s(buf, L"SetFontTextObject: sel=%d name=%ls isSystem=%d path=%ls",
			g_selectedFontIndex, item.displayName.c_str(), item.isSystemFont ? 1 : 0, item.filePath.c_str());
		logger->log(logger, buf);
	}
	if (!edit_handle)
	{
		if (logger)
			logger->error(logger, L"編集ハンドルが利用できません");
		return false;
	}

	struct SetObjectParam
	{
		std::string sysNameUtf8;
		std::string displayNameUtf8;
		std::string filePathUtf8;
		bool isSystem = false;
		bool updated = false;
	};

	SetObjectParam param{};
	param.isSystem = item.isSystemFont;
	param.sysNameUtf8 = ToUtf8(item.displayName);
	{
		std::wstring family = item.displayName;
		if (!item.isSystemFont)
		{
			size_t pos = family.find(L" [");
			if (pos != std::wstring::npos)
				family = family.substr(0, pos);
		}
		param.displayNameUtf8 = ToUtf8(family);
	}
	param.filePathUtf8 = ToUtf8(item.filePath);

	bool called = edit_handle->call_edit_section_param(&param, [](void *p, EDIT_SECTION *edit)
																   {
		auto *ctx = static_cast<SetObjectParam *>(p);
		if (!ctx || !edit || !edit->set_object_item_value)
			return;

		auto applyToObject = [&](OBJECT_HANDLE object) -> bool
		{
			if (!object)
				return false;
			bool ok = false;

			// 1) Variable Font Text (this plugin)
			if (ctx->isSystem)
			{
				// 2) Standard Text object (best-effort)
				ok |= !!edit->set_object_item_value(object, L"テキスト", L"フォント", ctx->sysNameUtf8.c_str());
				ok |= !!edit->set_object_item_value(object, L"Variable Font Text", L"フォント", ctx->sysNameUtf8.c_str());
				ok |= !!edit->set_object_item_value(object, L"Variable Font Text", L"フォントファイル", "");
			}
			else
			{
				ok |= !!edit->set_object_item_value(object, L"Variable Font Text", L"フォント", "");
				ok |= !!edit->set_object_item_value(object, L"Variable Font Text", L"フォントファイル", ctx->filePathUtf8.c_str());
			}
			return ok;
		};

		int updatedCount = 0;
		int n = edit->get_selected_object_num ? edit->get_selected_object_num() : 0;
		if (n >= 1)
		{
			for (int i = 0; i < n; i++)
			{
				OBJECT_HANDLE object = edit->get_selected_object(i);
				if (applyToObject(object))
					updatedCount++;
			}
			ctx->updated = (updatedCount > 0);
			return;
		}

		OBJECT_HANDLE object = edit->get_focus_object();
		ctx->updated = applyToObject(object);
	});

	return called && param.updated;
}

//---------------------------------------------------------------------
//	Layout and UI helpers
//---------------------------------------------------------------------
void UpdateLayout(HWND hwnd)
{
	if (!hwnd)
		return;
	RECT rc;
	GetClientRect(hwnd, &rc);
	int w = rc.right - rc.left;
	int h = rc.bottom - rc.top;
	int margin = 10;
	int searchHeight = 28;
	int nameHeight = 26;
	int typeHeight = 22;
	int detailGap = 8;
	int sampleRowHeight = 30;
	int listMinHeight = 160;
	int bottomMinHeight = 220;
	int buttonHeight = 28;

	int y = margin;
	if (g_hwndNameLabel)
		MoveWindow(g_hwndNameLabel, margin, y, w - margin * 2, nameHeight, TRUE);
	y += nameHeight + margin;
	if (g_hwndSearch)
		MoveWindow(g_hwndSearch, margin, y, w - margin * 2 - 140, searchHeight, TRUE);
	if (g_hwndType)
		MoveWindow(g_hwndType, w - margin - 130, y, 130, searchHeight, TRUE);
	y += searchHeight + margin;

	int actionY = y;
	int actionRight = w - margin;
	int textBtnW = 70;
	int vfBtnW = 120;
	int actionGap = 8;
	int labelW = std::max(80, actionRight - margin - (vfBtnW + actionGap + textBtnW + actionGap));
	if (g_hwndTypeLabel)
		MoveWindow(g_hwndTypeLabel, margin, actionY, labelW, typeHeight, TRUE);
	if (g_hwndAddText)
		MoveWindow(g_hwndAddText, actionRight - textBtnW, actionY - 2, textBtnW, buttonHeight, TRUE);
	if (g_hwndAddVF)
		MoveWindow(g_hwndAddVF, actionRight - textBtnW - actionGap - vfBtnW, actionY - 2, vfBtnW, buttonHeight, TRUE);
	y += typeHeight + detailGap;

	int remaining = h - y - margin;
	int listHeight = std::max(listMinHeight, remaining / 2);
	int bottomHeight = std::max(bottomMinHeight, remaining - listHeight);
	if (listHeight + bottomHeight > remaining)
		bottomHeight = remaining - listHeight;
	if (bottomHeight < bottomMinHeight)
		bottomHeight = bottomMinHeight;
	if (g_hwndGrid)
		MoveWindow(g_hwndGrid, margin, y, w - margin * 2, listHeight, TRUE);
	if (g_hwndGrid)
	{
		int gridW = w - margin * 2;
		ListView_SetColumnWidth(g_hwndGrid, 0, gridW - 4);
	}
	y += listHeight + margin;

	int paneHeight = bottomHeight - sampleRowHeight - margin;
	if (paneHeight < 120)
		paneHeight = 120;
	int previewW = (w - margin * 3) * 2 / 3;
	int axisW = w - margin * 3 - previewW;
	if (g_hwndPreview)
		MoveWindow(g_hwndPreview, margin, y, previewW, paneHeight, TRUE);
	if (g_hwndAxisLabel)
		MoveWindow(g_hwndAxisLabel, margin + previewW + margin, y, axisW, paneHeight, TRUE);

	int sampleTop = y + paneHeight + margin;
	if (g_hwndSample)
	{
		int bgW = 110;
		int sampleW = std::max(80, w - margin * 3 - bgW);
		MoveWindow(g_hwndSample, margin, sampleTop, sampleW, sampleRowHeight, TRUE);
	}
	if (g_hwndBgBtn)
		MoveWindow(g_hwndBgBtn, w - margin - 110, sampleTop, 110, sampleRowHeight, TRUE);
}

// Create child windows (labels, edit boxes, combo, listview, preview area,
// and buttons). This central helper encapsulates creation order and the
// initial control properties so RegisterPlugin remains concise.
static void CreateControls(HWND hwnd)
{
	g_hwndNameLabel = CreateWindowExW(0, WC_STATIC, L"", WS_VISIBLE | WS_CHILD | SS_LEFT,
									  10, 10, 400, 24, hwnd, (HMENU)IDC_NAME_LABEL, GetModuleHandleW(nullptr), nullptr);
	g_hwndSearch = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, L"", WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL,
								   10, 40, 400, 24, hwnd, (HMENU)IDC_SEARCH_EDIT, GetModuleHandleW(nullptr), nullptr);
	g_hwndType = CreateWindowExW(0, WC_COMBOBOX, nullptr, WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST,
								 420, 40, 200, 120, hwnd, (HMENU)IDC_TYPE_FILTER, GetModuleHandleW(nullptr), nullptr);
	SendMessageW(g_hwndType, CB_ADDSTRING, 0, (LPARAM)L"すべて");
	SendMessageW(g_hwndType, CB_ADDSTRING, 0, (LPARAM)L"システム");
	SendMessageW(g_hwndType, CB_ADDSTRING, 0, (LPARAM)L"外部");
	SendMessageW(g_hwndType, CB_SETCURSEL, 0, 0);

	g_hwndTypeLabel = CreateWindowExW(0, WC_STATIC, L"", WS_VISIBLE | WS_CHILD | SS_LEFT,
									  10, 70, 200, 24, hwnd, (HMENU)IDC_TYPE_LABEL, GetModuleHandleW(nullptr), nullptr);
	g_hwndAddVF = CreateWindowExW(0, WC_BUTTON, L"VF＋", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
								250, 70, 120, 28, hwnd, (HMENU)IDC_ADD_VF_BUTTON, GetModuleHandleW(nullptr), nullptr);
	g_hwndAddText = CreateWindowExW(0, WC_BUTTON, L"＋", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
								500, 70, 70, 28, hwnd, (HMENU)IDC_ADD_BUTTON, GetModuleHandleW(nullptr), nullptr);
	g_hwndAxisLabel = CreateWindowExW(WS_EX_CLIENTEDGE, WC_STATIC, L"", WS_VISIBLE | WS_CHILD | SS_LEFT,
									  10, 100, 400, 80, hwnd, (HMENU)IDC_AXIS_LABEL, GetModuleHandleW(nullptr), nullptr);
	g_hwndPreview = CreateWindowExW(WS_EX_CLIENTEDGE, WC_STATIC, L"", WS_VISIBLE | WS_CHILD,
									10, 190, 400, 200, hwnd, nullptr, GetModuleHandleW(nullptr), nullptr);

	g_hwndGrid = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"", WS_VISIBLE | WS_CHILD | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | LVS_NOCOLUMNHEADER,
								 10, 200, 600, 360, hwnd, (HMENU)IDC_FONT_GRID, GetModuleHandleW(nullptr), nullptr);
	if (g_hwndGrid)
	{
		ListView_SetExtendedListViewStyle(g_hwndGrid, LVS_EX_DOUBLEBUFFER | LVS_EX_FULLROWSELECT | LVS_EX_ONECLICKACTIVATE);
		LVCOLUMNW col{};
		col.mask = LVCF_WIDTH;
		col.cx = 580;
		ListView_InsertColumn(g_hwndGrid, 0, &col);
		if (logger)
			logger->info(logger, L"Created ListView grid for fonts");
	}
	else
	{
		if (logger)
			logger->error(logger, L"Failed to create ListView for font grid");
	}

	g_hwndSample = CreateWindowExW(WS_EX_CLIENTEDGE, WC_EDIT, g_sampleText.c_str(), WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL,
								   10, 580, 320, 24, hwnd, (HMENU)IDC_SAMPLE_TEXT_EDIT, GetModuleHandleW(nullptr), nullptr);
	g_hwndBgBtn = CreateWindowExW(0, WC_BUTTON, L"背景色", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
								  340, 580, 200, 28, hwnd, (HMENU)IDC_BG_COLOR_BTN, GetModuleHandleW(nullptr), nullptr);
}

// handle selection/change/click/dblclk for the font list
static void HandleListViewSelection(HWND hwnd, int hintIdx, bool dblclk) {
	int idx = ListView_GetNextItem(g_hwndGrid, -1, LVNI_SELECTED);
	if (idx < 0) idx = hintIdx;
	if (idx < 0) return;
	LVITEMW lvi{};
	lvi.iItem = idx;
	lvi.iSubItem = 0;
	lvi.mask = LVIF_PARAM;
	if (!ListView_GetItem(g_hwndGrid, &lvi)) return;

	int fontIdx = (int)lvi.lParam;
	if (fontIdx < 0 || fontIdx >= (int)g_fontList.size()) return;
	if (fontIdx == g_selectedFontIndex && !dblclk)
	{
		if (logger)
			logger->verbose(logger, L"ListView: selection unchanged, skip");
		return;
	}
	g_selectedFontIndex = fontIdx;
	UpdateDetailPanel();
	RedrawGrid();
	RenderPreview(L"ListViewSelection");
	if (logger) {
		wchar_t buf[128];
		swprintf_s(buf, L"ListView: select idx=%d fontIndex=%d", idx, g_selectedFontIndex);
		logger->info(logger, buf);
	}
	if (dblclk) {
		PostMessageW(hwnd, WM_DO_SET_FONT_OBJECT, 0, 0);
	}
}

// background color chooser handler
static void HandleBgColorButton(HWND hwnd) {
	CHOOSECOLORW cc{};
	static COLORREF cust[16] = {0};
	cc.lStructSize = sizeof(cc);
	cc.hwndOwner = hwnd;
	cc.lpCustColors = cust;
	cc.rgbResult = g_previewBgColor;
	cc.Flags = CC_FULLOPEN | CC_RGBINIT;
	if (ChooseColorW(&cc)) {
		g_previewBgColor = cc.rgbResult;
		RedrawGrid();
		RenderPreview(L"BgColorChange");
	}
}

void ApplyFilterFromUI()
{
	wchar_t buf[512];
	if (g_hwndSearch)
	{
		GetWindowTextW(g_hwndSearch, buf, _countof(buf));
		g_searchQuery = buf;
	}
	if (g_hwndType)
	{
		int sel = (int)SendMessageW(g_hwndType, CB_GETCURSEL, 0, 0);
		if (sel == 1)
			g_filterType = FontTypeFilter::System;
		else if (sel == 2)
			g_filterType = FontTypeFilter::Folder;
		else
			g_filterType = FontTypeFilter::All;
	}
	ApplyFilter();
}

// forward decls for refactor helpers
static void CreateControls(HWND hwnd);

//---------------------------------------------------------------------
//	Window procedure
//---------------------------------------------------------------------
LRESULT CALLBACK wnd_proc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam)
{
	switch (message)
	{
	case WM_LBUTTONDOWN:
		// Allow default processing; ListView handles its own clicks
		break;
	case WM_DO_SET_FONT_OBJECT:
		// Apply selected font to the host object.
		if (!SetFontTextObject()) {
			if (logger) logger->warn(logger, L"SetFontTextObject reported failure or no selection");
		}
		return 0;
	case WM_MOUSEWHEEL:
		// ListView manages scrolling; no parent fallback needed
		break;
	case WM_COMMAND:
		switch (LOWORD(wparam))
		{
		case IDC_SEARCH_EDIT:
			if (HIWORD(wparam) == EN_CHANGE)
				ApplyFilterFromUI();
			return 0;
		case IDC_TYPE_FILTER:
			if (HIWORD(wparam) == CBN_SELCHANGE)
				ApplyFilterFromUI();
			return 0;
		case IDC_SAMPLE_TEXT_EDIT:
			if (HIWORD(wparam) == EN_CHANGE)
			{
				SyncSampleTextFromEdit();
				RedrawGrid();
				RenderPreview(L"SampleTextChange");
			}
			return 0;
			case IDC_BG_COLOR_BTN:
			{
				if (HIWORD(wparam) == BN_CLICKED) {
					HandleBgColorButton(hwnd);
				}
				return 0;
			}
		case IDC_ADD_VF_BUTTON:
			if (HIWORD(wparam) == BN_CLICKED)
				CreateVariableFontObject(LOWORD(wparam));
			return 0;
		case IDC_ADD_BUTTON:
			if (HIWORD(wparam) == BN_CLICKED)
				CreateVariableFontObject(LOWORD(wparam));
			return 0;
		}
		break;
	case WM_NOTIFY:
	{
		LPNMHDR pnm = (LPNMHDR)lparam;
		if (pnm->idFrom == IDC_FONT_GRID)
		{
			if (pnm->code == LVN_ITEMACTIVATE || pnm->code == NM_CLICK || pnm->code == LVN_ITEMCHANGED || pnm->code == NM_DBLCLK)
			{
				int hint = -1;
				if (pnm->code == LVN_ITEMACTIVATE || pnm->code == NM_CLICK || pnm->code == NM_DBLCLK)
				{
					LPNMITEMACTIVATE act = (LPNMITEMACTIVATE)lparam;
					hint = act ? act->iItem : -1;
				}
				else if (pnm->code == LVN_ITEMCHANGED)
				{
					LPNMLISTVIEW lvn = (LPNMLISTVIEW)lparam;
					if ((lvn->uNewState & LVIS_SELECTED) != 0) hint = lvn->iItem;
				}

				if (hint >= 0) {
					HandleListViewSelection(hwnd, hint, pnm->code == NM_DBLCLK);
				}
			}
		}
		break;
	}
	case WM_PAINT:
		RenderPreview(L"WM_PAINT");
		break;
	case WM_SIZE:
		UpdateLayout(hwnd);
		RenderPreview(L"WM_SIZE");
		return 0;
	}
	return DefWindowProc(hwnd, message, wparam, lparam);
}

//---------------------------------------------------------------------
//	Logger initialization
//---------------------------------------------------------------------
EXTERN_C __declspec(dllexport) void InitializeLogger(LOG_HANDLE *handle)
{
	logger = handle;
}

//---------------------------------------------------------------------
//	Plugin DLL initialization
//---------------------------------------------------------------------
EXTERN_C __declspec(dllexport) bool InitializePlugin(DWORD)
{
	return SUCCEEDED(InitializeGraphics());
}

//---------------------------------------------------------------------
//	Plugin DLL cleanup
//---------------------------------------------------------------------
EXTERN_C __declspec(dllexport) void UninitializePlugin()
{
	g_dwriteFactory.Reset();
	g_d2dTarget.Reset();
	g_d2dContext.Reset();
	g_d2dDevice.Reset();
	g_d2dFactory.Reset();
	g_swapChain.Reset();
	g_d3dContext.Reset();
	g_d3dDevice.Reset();
}

//---------------------------------------------------------------------
//	Plugin registration
//---------------------------------------------------------------------
EXTERN_C __declspec(dllexport) void RegisterPlugin(HOST_APP_TABLE *host)
{
	host->set_plugin_information(L"Font Preview Grid v1.0 By 黒猫大福");

	INITCOMMONCONTROLSEX icc{};
	icc.dwSize = sizeof(icc);
	icc.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&icc);

	WNDCLASSEXW wcex{};
	wcex.cbSize = sizeof(WNDCLASSEXW);
	wcex.lpszClassName = FontPreviewWindowName;
	wcex.lpfnWndProc = wnd_proc;
	wcex.hInstance = GetModuleHandleW(nullptr);
	wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
	if (!RegisterClassExW(&wcex))
		return;

	HWND hwnd = CreateWindowExW(0, FontPreviewWindowName, FontPreviewWindowName, WS_POPUP,
								CW_USEDEFAULT, CW_USEDEFAULT, 640, 720, nullptr, nullptr, GetModuleHandleW(0), nullptr);
	if (!hwnd)
		return;

	// Let DWM decide non-client rendering based on window style/theme
	// Use DWMNCRP_USEWINDOWSTYLE so the system matches the current theme.
	// This helps non-client areas (caption/border) blend with system theme.
	{
		DWMNCRENDERINGPOLICY policy = DWMNCRP_USEWINDOWSTYLE;
		HRESULT hr = DwmSetWindowAttribute(hwnd, DWMWA_NCRENDERING_POLICY, &policy, sizeof(policy));
		if (FAILED(hr) && logger)
		{
			wchar_t buf[128];
			swprintf_s(buf, L"DwmSetWindowAttribute(DWMWA_NCRENDERING_POLICY) failed 0x%08x", hr);
			logger->warn(logger, buf);
		}
	}

	CreateControls(hwnd);

	EnumerateFonts();
	ApplyFilter();
	RebuildListViewItems();
	UpdateDetailPanel();
	UpdateLayout(hwnd);
	RenderPreview(L"RegisterPlugin init");
	ShowWindow(hwnd, SW_SHOW);
	UpdateWindow(hwnd);

	host->register_window_client(FontPreviewWindowName, hwnd);
	edit_handle = host->create_edit_handle();
}
