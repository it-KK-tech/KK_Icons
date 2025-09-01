/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Streamline API Configuration
interface StreamlineConfig {
  apiKey?: string; // optional in dev when proxy injects header
  baseUrl: string;
}

// Streamline API Response Types
interface StreamlineIcon {
  hash: string;
  name: string;
  imagePreviewUrl: string;
  isFree: boolean;
  familySlug: string;
  familyName: string;
  categorySlug: string;
  categoryName: string;
  subcategorySlug: string;
  subcategoryName: string;
}

interface StreamlineSearchResponse {
  query: string;
  results: StreamlineIcon[];
  pagination: {
    total: number;
    hasMore: boolean;
    offset: number;
    nextOffset: number;
  };
}

// Internal Icon interface for UI
interface Icon {
  hash: string;
  name: string;
  family: string;
  category: string;
  svgUrl?: string;
  tags: string[];
}

// Configuration
const streamlineConfig: StreamlineConfig = {
  // Do not store secrets in client code. In dev, the webpack devServer proxy injects x-api-key
  // For production, configure your backend/proxy to add the header server-side.
  baseUrl: "/api/streamline"
};

let currentIcons: Icon[] = [];
let searchTimeout: number;
let currentPage = 1;
let isLoading = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    initializeIconFinder();
  }
});

// Streamline API Service Functions
class StreamlineAPIService {
  private static async makeRequest(endpoint: string, params: Record<string, string> = {}): Promise<any> {
    // In development, the devServer proxy injects the API key header.

    // Handle relative URLs for proxy
    let url: URL;
    if (streamlineConfig.baseUrl.startsWith('/')) {
      // For proxy paths like '/api/streamline', construct URL relative to current origin
      url = new URL(`${streamlineConfig.baseUrl}${endpoint}`, window.location.origin);
    } else {
      // For absolute URLs like 'https://...'
      url = new URL(`${streamlineConfig.baseUrl}${endpoint}`);
    }
    
    Object.entries(params).forEach(([key, value]) => {
      if (value) url.searchParams.append(key, value);
    });

    try {
      const response = await fetch(url.toString(), {
        method: 'GET',
        headers: {
          // Do not attach client-side secret; rely on proxy
          'accept': 'application/json'
        }
      });

      if (!response.ok) {
        if (response.status === 401) {
          throw new Error("Invalid API key. Please check your Streamline API credentials.");
        } else if (response.status === 429) {
          throw new Error("Rate limit exceeded. Please wait a moment before searching again.");
        } else {
          throw new Error(`API request failed: ${response.status} ${response.statusText}`);
        }
      }

      return await response.json();
    } catch (error) {
      console.error("Fetch error details:", error);
      
      // Check if it's a network/CORS error
      if (error instanceof TypeError) {
        if (error.message.includes('fetch') || error.message.includes('NetworkError') || error.message.includes('Failed to fetch')) {
          throw new Error(`CORS/Network error: The Streamline API is blocking requests from Office Add-ins. This is a common issue where the API server doesn't allow cross-origin requests from Office domains. URL attempted: ${url.toString()}`);
        }
      }
      
      // Re-throw other errors as-is
      throw error;
    }
  }

  static async searchFamily(familySlug: string, query?: string, page: number = 1, perPage: number = 100): Promise<StreamlineSearchResponse> {
    const params: Record<string, string> = {
      page: page.toString(),
      per_page: perPage.toString()
    };

    if (query) {
      params.query = query;
    }

    return await this.makeRequest(`/search/family/${familySlug}`, params);
  }

  static async downloadIconSVG(iconHash: string): Promise<string> {

    // Use proxy for downloads to avoid CORS issues
    let downloadUrl: string;
    if (streamlineConfig.baseUrl.startsWith('/')) {
      // For proxy paths like '/api/streamline', construct URL relative to current origin
      downloadUrl = `${window.location.origin}${streamlineConfig.baseUrl}/icons/${iconHash}/download/svg?size=48&responsive=false`;
    } else {
      // For absolute URLs like 'https://...'
      downloadUrl = `${streamlineConfig.baseUrl}/icons/${iconHash}/download/svg?size=48&responsive=false`;
    }

    const response = await fetch(downloadUrl, {
      headers: {
        // rely on proxy to inject the key
        'accept': 'image/svg+xml'
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to download SVG: ${response.status}`);
    }

    return await response.text();
  }
}

function initializeIconFinder() {
  const searchInput = document.getElementById("icon-search") as HTMLInputElement;

  // Show initial message
  updateResultsCount(0, "");
  showApiKeyMessage();

  // Add search event listener with debouncing
  searchInput.addEventListener("input", () => {
    clearTimeout(searchTimeout);
    searchTimeout = window.setTimeout(() => {
      performSearch();
    }, 500);
  });

  // Add enter key support for search
  searchInput.addEventListener("keypress", (e) => {
    if (e.key === "Enter") {
      clearTimeout(searchTimeout);
      performSearch();
    }
  });
}

function showApiKeyMessage() {
  // No special message; in dev the proxy handles the key. If requests fail with 401,
  // developers should set STREAMLINE_API_KEY in .env for the dev server.
  const iconsGrid = document.getElementById("icons-grid");
  iconsGrid.innerHTML = `
    <div style="grid-column: 1 / -1; text-align: center; padding: 40px; color: #605e5c;">
      <div style="font-size: 16px;">Type a search term to find icons</div>
    </div>
  `;
}

async function performSearch() {
  if (isLoading) return;

  const searchInput = document.getElementById("icon-search") as HTMLInputElement;
  const loadingSpinner = document.getElementById("loading-spinner");
  const iconsGrid = document.getElementById("icons-grid");

  const searchTerm = searchInput.value.trim();

  // Don't search with empty term
  if (!searchTerm) {
    showApiKeyMessage();
    return;
  }

  isLoading = true;
  currentPage = 1;

  // Show loading spinner
  loadingSpinner.style.display = "flex";
  iconsGrid.style.display = "none";

  try {
    // No client-side key check; rely on proxy

    // Search within streamline-light family
    const response = await StreamlineAPIService.searchFamily('streamline-light', searchTerm, currentPage);

    // Convert Streamline icons to internal format
    currentIcons = response.results.map(streamlineIcon => ({
      hash: streamlineIcon.hash,
      name: streamlineIcon.name,
      family: streamlineIcon.familySlug,
      category: streamlineIcon.categoryName || streamlineIcon.familyName,
      tags: [],
      svgUrl: streamlineIcon.imagePreviewUrl
    }));

    displayIcons(currentIcons);
    updateResultsCount(response.pagination.total, undefined, Math.floor(response.pagination.offset / 100) + 1, Math.ceil(response.pagination.total / 100));

  } catch (error) {
    console.error("Search error:", error);
    showErrorMessage(error.message || "Failed to search icons. Please try again.");
  } finally {
    isLoading = false;
    loadingSpinner.style.display = "none";
    iconsGrid.style.display = "grid";
  }
}

function displayIcons(icons: Icon[]) {
  const iconsGrid = document.getElementById("icons-grid");

  if (icons.length === 0) {
    iconsGrid.innerHTML = `
      <div style="grid-column: 1 / -1; text-align: center; padding: 40px; color: #605e5c;">
        <div style="font-size: 48px; margin-bottom: 16px;">🔍</div>
        <div style="font-size: 16px; margin-bottom: 8px;">No icons found</div>
        <div style="font-size: 14px;">Try different search terms or browse icon families</div>
      </div>
    `;
    return;
  }

  iconsGrid.innerHTML = icons.map(icon => `
    <div class="icon-item" onclick="insertStreamlineIcon('${icon.hash}', '${icon.name}')" title="Click to copy ${icon.name} to clipboard">
      <div class="icon-symbol" data-icon-hash="${icon.hash}">
        ${icon.svgUrl ? 
          `<img src="${icon.svgUrl}" alt="${icon.name}" />` : 
          '<div class="icon-placeholder">📄</div>'
        }
      </div>
    </div>
  `).join("");
}

function showErrorMessage(message: string) {
  const iconsGrid = document.getElementById("icons-grid");
  iconsGrid.innerHTML = `
    <div style="grid-column: 1 / -1; text-align: center; padding: 40px; color: #d13438;">
      <div style="font-size: 48px; margin-bottom: 16px;">⚠️</div>
      <div style="font-size: 16px; margin-bottom: 8px; font-weight: 600;">Error</div>
      <div style="font-size: 14px; max-width: 300px; margin: 0 auto; line-height: 1.4;">${message}</div>
    </div>
  `;
}

function updateResultsCount(count: number, customMessage?: string, page?: number, totalPages?: number) {
  const resultsCount = document.getElementById("results-count");

  if (customMessage) {
    resultsCount.textContent = customMessage;
    return;
  }

  if (count === 0) {
    resultsCount.textContent = "No icons found";
  } else if (count === 1) {
    resultsCount.textContent = "1 icon found";
  } else {
    let message = `${count.toLocaleString()} icons found`;
    if (page && totalPages) {
      message += ` (Page ${page} of ${totalPages})`;
    }
    resultsCount.textContent = message;
  }
}

// Make functions globally available for onclick handlers
(window as any).insertStreamlineIcon = insertStreamlineIcon;

async function insertStreamlineIcon(iconHash: string, iconName: string) {
  try {
    // Download the SVG content
    const svgContent = await StreamlineAPIService.downloadIconSVG(iconHash);
    
    // Copy the raw SVG content to clipboard as both text and SVG
    try {
      // Create an SVG blob
      const svgBlob = new Blob([svgContent], { type: 'image/svg+xml' });
      
      // Copy to clipboard with multiple formats for maximum compatibility
      await navigator.clipboard.write([
        new ClipboardItem({
          'image/svg+xml': svgBlob,
          'text/plain': new Blob([svgContent], { type: 'text/plain' })
        })
      ]);
      
      showSuccessNotification(`${iconName} SVG icon copied to clipboard! Press Ctrl+V to paste in PowerPoint.`);
      
    } catch (clipboardError) {
      console.error("Advanced clipboard copy failed, trying simple text copy:", clipboardError);
      
      // Fallback: Copy as text only
      await navigator.clipboard.writeText(svgContent);
      showSuccessNotification(`${iconName} SVG copied as text to clipboard! Press Ctrl+V to paste.`);
    }
  } catch (error) {
    console.error("Error copying icon:", error);
    showErrorNotification(error.message || "Failed to copy icon. Please try again.");
  }
}

function showSuccessNotification(message: string) {
  const notification = document.createElement("div");
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    background: #107c10;
    color: white;
    padding: 12px 16px;
    border-radius: 4px;
    font-size: 14px;
    z-index: 1000;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    animation: slideIn 0.3s ease-out;
  `;
  notification.textContent = message;

  // Add animation styles
  const style = document.createElement('style');
  style.textContent = `
    @keyframes slideIn {
      from { transform: translateX(100%); opacity: 0; }
      to { transform: translateX(0); opacity: 1; }
    }
  `;
  document.head.appendChild(style);

  document.body.appendChild(notification);

  setTimeout(() => {
    if (notification.parentNode) {
      document.body.removeChild(notification);
    }
    if (style.parentNode) {
      document.head.removeChild(style);
    }
  }, 3000);
}

function showErrorNotification(message: string) {
  const notification = document.createElement("div");
  notification.style.cssText = `
    position: fixed;
    top: 20px;
    right: 20px;
    background: #d13438;
    color: white;
    padding: 12px 16px;
    border-radius: 4px;
    font-size: 14px;
    z-index: 1000;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    animation: slideIn 0.3s ease-out;
  `;
  notification.textContent = message;

  document.body.appendChild(notification);

  setTimeout(() => {
    if (notification.parentNode) {
      document.body.removeChild(notification);
    }
  }, 4000);
}

export { insertStreamlineIcon };