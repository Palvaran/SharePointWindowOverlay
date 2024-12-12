import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { override } from "@microsoft/decorators";
import config from "./config.json";

import * as strings from "TobyOverlayCustomizerApplicationCustomizerStrings";

const LOG_SOURCE: string = "TobyOverlayCustomizerApplicationCustomizer";

/**
 * Define the properties for your Application Customizer if needed.
 */
export interface ITobyOverlayCustomizerApplicationCustomizerProperties {
  // You can define custom properties here if needed
  // For example, iframeUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class TobyOverlayCustomizerApplicationCustomizer extends BaseApplicationCustomizer<ITobyOverlayCustomizerApplicationCustomizerProperties> {
  private iframeContainer: HTMLDivElement;
  private iframe: HTMLIFrameElement;
  private minimizeButton: HTMLButtonElement;
  private restoreButton: HTMLButtonElement;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Initialize UI Elements
    this.initializeIframeContainer();
    this.initializeIframe();
    this.initializeMinimizeButton();
    this.initializeRestoreButton();
    this.addEventListeners();
    this.handleInitialState();

    // Optional: Log that the iframe has been added
    Log.info(LOG_SOURCE, "Iframe Customizer has been initialized.");

    return Promise.resolve();
  }

  /**
   * Initializes the container for the iframe and buttons.
   */
  private initializeIframeContainer(): void {
    this.iframeContainer = document.createElement("div");
    this.iframeContainer.id = "customIframeContainer";
    this.iframeContainer.style.position = "fixed";
    this.iframeContainer.style.bottom = "20px";
    this.iframeContainer.style.right = "20px";
    this.iframeContainer.style.width = "300px";
    this.iframeContainer.style.height = "500px";
    this.iframeContainer.style.zIndex = "1000";
    this.iframeContainer.style.background = "white";
    this.iframeContainer.style.boxShadow = "0 4px 8px rgba(0, 0, 0, 0.2)";
    this.iframeContainer.style.borderRadius = "8px";
    this.iframeContainer.style.overflow = "hidden";
    this.iframeContainer.style.transition = "all 0.3s ease";
    this.iframeContainer.style.display = "flex";
    this.iframeContainer.style.flexDirection = "column";
    this.iframeContainer.style.alignItems = "flex-end";
    this.iframeContainer.style.pointerEvents = "auto"; // Ensure interactions are possible

    // Append the container to the body
    document.body.appendChild(this.iframeContainer);
  }

  /**
   * Initializes the iframe element with lazy loading.
   */
  private initializeIframe(): void {
    this.iframe = document.createElement("iframe");
    this.iframe.id = "customIframe";
    this.iframe.style.flex = "1"; // Allow the iframe to take up remaining space
    this.iframe.style.width = "100%";
    this.iframe.style.border = "none";
    this.iframe.style.display = "none"; // Initially hidden for lazy loading
    this.iframe.title = "Toby Overlay"; // Accessibility: Provide a title

    // Add error handling for iframe load failure
    this.iframe.onerror = () => {
      Log.error(LOG_SOURCE, new Error("Failed to load the iframe."));
      this.displayErrorMessage();
    };

    // Append the iframe to the container
    this.iframeContainer.appendChild(this.iframe);
  }

  /**
   * Initializes the minimize button.
   */
  private initializeMinimizeButton(): void {
    this.minimizeButton = document.createElement("button");
    this.minimizeButton.innerText = "−"; // Minimize symbol
    this.minimizeButton.style.background = "#333";
    this.minimizeButton.style.color = "white";
    this.minimizeButton.style.border = "none";
    this.minimizeButton.style.borderRadius = "50%";
    this.minimizeButton.style.width = "30px";
    this.minimizeButton.style.height = "30px";
    this.minimizeButton.style.cursor = "pointer";
    this.minimizeButton.style.zIndex = "1001";
    this.minimizeButton.style.fontSize = "18px";
    this.minimizeButton.style.lineHeight = "30px";
    this.minimizeButton.style.textAlign = "center";
    this.minimizeButton.title = "Minimize";
    this.minimizeButton.setAttribute("aria-label", "Minimize iframe");
    this.minimizeButton.setAttribute("tabindex", "0");

    // Initially visible
    this.minimizeButton.style.display = "block";

    // Append the minimize button to the container
    this.iframeContainer.appendChild(this.minimizeButton);
  }

  /**
   * Initializes the restore button (hidden by default).
   */
  private initializeRestoreButton(): void {
    this.restoreButton = document.createElement("button");
    this.restoreButton.innerText = "🡅"; // Restore symbol
    this.restoreButton.style.position = "fixed";
    this.restoreButton.style.bottom = "20px";
    this.restoreButton.style.right = "20px";
    this.restoreButton.style.width = "30px";
    this.restoreButton.style.height = "30px";
    this.restoreButton.style.background = "#333";
    this.restoreButton.style.color = "white";
    this.restoreButton.style.border = "none";
    this.restoreButton.style.borderRadius = "50%";
    this.restoreButton.style.cursor = "pointer";
    this.restoreButton.style.zIndex = "1001";
    this.restoreButton.style.fontSize = "18px";
    this.restoreButton.style.lineHeight = "30px";
    this.restoreButton.style.textAlign = "center";
    this.restoreButton.title = "Restore";
    this.restoreButton.setAttribute("aria-label", "Restore iframe");
    this.restoreButton.setAttribute("tabindex", "0");
    this.restoreButton.style.display = "none"; // Hidden initially

    // Append the restore button to the body
    document.body.appendChild(this.restoreButton);
  }

  /**
   * Adds event listeners for buttons and window resize.
   */
  private addEventListeners(): void {
    // Minimize button click event
    this.minimizeButton.addEventListener("click", () => {
      this.minimizeIframe();
    });

    // Restore button click event
    this.restoreButton.addEventListener("click", () => {
      this.restoreIframe();
    });

    // Keyboard accessibility for minimize button
    this.minimizeButton.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        this.minimizeIframe();
      }
    });

    // Keyboard accessibility for restore button
    this.restoreButton.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        this.restoreIframe();
      }
    });

    // Lazy load the iframe after a delay or based on user interaction
    window.addEventListener("load", () => {
      this.lazyLoadIframe();
    });

    // Adjust iframe size on window resize
    window.addEventListener("resize", () => {
      this.adjustIframeSize();
    });
  }

  /**
   * Handles the initial state of the iframe based on session storage.
   */
  private handleInitialState(): void {
    const isMinimized = sessionStorage.getItem("iframeMinimized") === "true";

    if (isMinimized) {
      this.minimizeIframe(false); // Minimize without saving state again
    } else {
      this.restoreIframe(false); // Ensure iframe is expanded
    }
  }

  /**
   * Minimizes the iframe container.
   * @param saveState Whether to save the state in session storage.
   */
  private minimizeIframe(saveState: boolean = true): void {
    this.iframeContainer.style.height = "50px"; // Shrink height
    this.iframe.style.display = "none"; // Hide iframe content
    this.minimizeButton.style.display = "none"; // Hide minimize button
    this.restoreButton.style.display = "block"; // Show restore button

    if (saveState) {
      sessionStorage.setItem("iframeMinimized", "true");
    }
  }

  /**
   * Restores the iframe container to its original size.
   * @param saveState Whether to save the state in session storage.
   */
  private restoreIframe(saveState: boolean = true): void {
    this.iframeContainer.style.height = "500px"; // Restore height
    this.iframe.style.display = "block"; // Show iframe content
    this.minimizeButton.style.display = "block"; // Show minimize button
    this.restoreButton.style.display = "none"; // Hide restore button

    if (saveState) {
      sessionStorage.setItem("iframeMinimized", "false");
    }
  }

  /**
   * Lazy loads the iframe by setting its `src` after a delay.
   */
  private lazyLoadIframe(): void {
    setTimeout(() => {
      this.iframe.src = config.iframeUrl; // Points to config.json
      this.iframe.style.display = "block"; // Display iframe after loading
    }, 1000); // Delay of 1 second
  }

  /**
   * Displays an error message within the iframe container if loading fails.
   */
  private displayErrorMessage(): void {
    const errorMessage: HTMLDivElement = document.createElement("div");
    errorMessage.innerText = "Failed to load the iframe.";
    errorMessage.style.color = "red";
    errorMessage.style.padding = "10px";
    errorMessage.style.textAlign = "center";
    errorMessage.style.width = "100%";
    errorMessage.style.boxSizing = "border-box";

    // Clear existing content and display the error message
    this.iframeContainer.innerHTML = "";
    this.iframeContainer.appendChild(errorMessage);
  }

  /**
   * Adjusts the iframe container size and position based on window size.
   */
  private adjustIframeSize(): void {
    if (window.innerWidth < 600) {
      // Mobile devices
      this.iframeContainer.style.width = "90%";
      this.iframeContainer.style.height = this.isMinimized() ? "50px" : "300px";
      this.iframeContainer.style.right = "5%";
      this.iframeContainer.style.bottom = "10px";
    } else if (window.innerWidth < 992) {
      // Tablets
      this.iframeContainer.style.width = "250px";
      this.iframeContainer.style.height = this.isMinimized() ? "50px" : "400px";
      this.iframeContainer.style.right = "10px";
      this.iframeContainer.style.bottom = "15px";
    } else {
      // Desktops
      this.iframeContainer.style.width = "300px";
      this.iframeContainer.style.height = this.isMinimized() ? "50px" : "500px";
      this.iframeContainer.style.right = "20px";
      this.iframeContainer.style.bottom = "20px";
    }
  }

  /**
   * Checks if the iframe is currently minimized.
   * @returns True if minimized, false otherwise.
   */
  private isMinimized(): boolean {
    return sessionStorage.getItem("iframeMinimized") === "true";
  }
}
