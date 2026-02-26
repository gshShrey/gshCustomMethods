/**
 * Base context available on all platforms
 */
export interface WalnutBaseContext {
  /** Platform discriminator - "web", "api", etc. */
  platform: 'web' | 'api' | 'shared';
  
  /** Base URL of the application under test */
  testBaseUrl: string;
  
  /** Resolved step arguments - values from ${...} placeholders in the step description, in order */
  args: any[];
  
  /** Shared variable store across steps (read/write) */
  variableContext: Record<string, any>;
  
  /** Resolve {{variable}} placeholders in a string */
  replacePlaceholders(text: string): string;
  
  /** Output a debug log message from the custom method */
  log(message: string): void;
  
  /** Output a warning log message from the custom method */
  warn(message: string): void;
  
  /** Store a value in the shared variable context for use across steps */
  setVariable(name: string, value: any): void;
  
  /** Retrieve a value from the shared variable context */
  getVariable(name: string): any;
}

/**
 * Browser automation context (Playwright-based)
 */
export interface WalnutWebContext extends WalnutBaseContext {
  platform: 'web';
  
  /** Raw Playwright Page instance for advanced operations */
  page: any;
  
  // Element actions
  click(selector: string): Promise<void>;
  type(selector: string, text: string): Promise<void>;
  fill(selector: string, text: string): Promise<void>;
  selectOption(selector: string, value: string): Promise<void>;
  hover(selector: string): Promise<void>;
  dblclick(selector: string): Promise<void>;
  focus(selector: string): Promise<void>;
  blur(selector: string): Promise<void>;
  pressKey(key: string): Promise<void>;
  drag(sourceSelector: string, targetSelector: string): Promise<void>;
  fileUpload(selector: string, filePath: string): Promise<void>;
  tap(selector: string): Promise<void>;
  check(selector: string): Promise<void>;
  uncheck(selector: string): Promise<void>;
  clear(selector: string): Promise<void>;
  
  // Navigation
  navigate(url: string): Promise<void>;
  navigateBack(): Promise<void>;
  reload(): Promise<void>;
  goForward(): Promise<void>;
  close(): Promise<void>;
  
  // Verification
  verifyElementVisible(selector: string): Promise<void>;
  verifyElementHidden(selector: string): Promise<void>;
  verifyTextVisible(text: string): Promise<void>;
  verifyValue(selector: string, expectedValue: string): Promise<void>;
  verifyUrlContains(substring: string): Promise<void>;
  verifyNavigation(urlPattern: string): Promise<void>;
  
  // Wait
  waitForVisible(selector: string, opts?: { timeout?: number }): Promise<void>;
  waitForHidden(selector: string, opts?: { timeout?: number }): Promise<void>;
  wait(ms: number): Promise<void>;
  waitForNavigation(opts?: { timeout?: number; waitUntil?: 'load' | 'domcontentloaded' | 'networkidle' }): Promise<void>;
  waitForUrl(urlPattern: string, opts?: { timeout?: number }): Promise<void>;
  waitForAttached(selector: string, opts?: { timeout?: number }): Promise<void>;
  waitForDetached(selector: string, opts?: { timeout?: number }): Promise<void>;
  
  // Query
  getText(selector: string): Promise<string>;
  getAttribute(selector: string, attribute: string): Promise<string | null>;
  getInputValue(selector: string): Promise<string>;
  getUrl(): Promise<string>;
  getTitle(): Promise<string>;
  count(selector: string): Promise<number>;
  isVisible(selector: string): Promise<boolean>;
  isEnabled(selector: string): Promise<boolean>;
  isChecked(selector: string): Promise<boolean>;
  
  // Mouse
  mouseClick(x: number, y: number): Promise<void>;
  mouseMove(x: number, y: number): Promise<void>;
  mouseDrag(fromX: number, fromY: number, toX: number, toY: number): Promise<void>;
  
  // Layout
  setViewportSize(width: number, height: number): Promise<void>;
  
  // Tab
  switchToTab(index: number): Promise<void>;
  getTabCount(): Promise<number>;
  
  // Advanced
  evaluate(script: string): Promise<any>;
  screenshot(opts?: { path?: string; fullPage?: boolean }): Promise<Buffer>;
  scroll(opts?: { x?: number; y?: number }): Promise<void>;
  handleDialog(action: 'accept' | 'dismiss', promptText?: string): Promise<void>;
}

/**
 * API response object
 */
export interface ApiResponse {
  /** HTTP status code */
  status: number;
  
  /** HTTP status text */
  statusText: string;
  
  /** Response headers */
  headers: Record<string, string>;
  
  /** Parsed response body (JSON object or string) */
  body: any;
  
  /** Response time in milliseconds */
  responseTime: number;
}

/**
 * Options for HTTP request methods
 */
export interface RequestOptions {
  /** Additional request headers */
  headers?: Record<string, string>;
  
  /** URL query parameters */
  params?: Record<string, string>;
  
  /** Request timeout in seconds */
  timeout?: number;
  
  /** Per-request authentication override */
  auth?: {
    type: 'bearer' | 'basic' | 'apiKey';
    token?: string;
    username?: string;
    password?: string;
    key?: string;
    value?: string;
    in?: 'header' | 'query';
  };
}

/**
 * HTTP request automation context
 */
export interface WalnutApiContext extends WalnutBaseContext {
  platform: 'api';
  
  // HTTP methods
  request(method: string, url: string, body?: any, opts?: RequestOptions): Promise<ApiResponse>;
  get(url: string, opts?: RequestOptions): Promise<ApiResponse>;
  post(url: string, body?: any, opts?: RequestOptions): Promise<ApiResponse>;
  put(url: string, body?: any, opts?: RequestOptions): Promise<ApiResponse>;
  patch(url: string, body?: any, opts?: RequestOptions): Promise<ApiResponse>;
  delete(url: string, opts?: RequestOptions): Promise<ApiResponse>;
  
  // State
  setHeader(name: string, value: string): void;
  setAuth(type: 'bearer' | 'basic' | 'apiKey', credentials: { token?: string; username?: string; password?: string; key?: string; value?: string; in?: 'header' | 'query' }): void;
  getCookies(url?: string): Promise<any[]>;
  setCookie(cookie: any, url?: string): Promise<void>;
  
  // Response
  extractFromBody(response: ApiResponse, jsonPath: string): any;
  
  // Assertions
  assertStatus(response: ApiResponse, expected: number): void;
  assertBodyContains(response: ApiResponse, text: string): void;
  assertResponseTime(response: ApiResponse, maxMs: number): void;
  assertHeader(response: ApiResponse, headerName: string, expected: string): void;
  assertBodyEquals(response: ApiResponse, jsonPath: string, expected: any): void;
  assertBodyMatches(response: ApiResponse, jsonPath: string, regex: RegExp): void;
}

/**
 * Shared context for data processing (no browser/HTTP)
 */
export interface WalnutSharedContext extends WalnutBaseContext {
  platform: 'shared';
}

/**
 * Union of all context types
 */
export type WalnutContext = WalnutWebContext | WalnutApiContext | WalnutSharedContext;
