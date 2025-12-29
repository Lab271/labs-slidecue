"use strict";
const electron = require("electron");
const path = require("path");
const os = require("os");
const fs = require("fs/promises");
const fsSync = require("fs");
const child_process = require("child_process");
const util = require("util");
const express = require("express");
const http = require("http");
const socket_io = require("socket.io");
const electronUpdater = require("electron-updater");
const execAsync$1 = util.promisify(child_process.exec);
async function parsePresentationData(pptxPath) {
  const slides = [];
  const visibleSlides = [];
  const hiddenSlides = [];
  try {
    const { stdout: fileList } = await execAsync$1(
      `unzip -l "${pptxPath}" 2>/dev/null | grep -E "ppt/slides/slide[0-9]+\\.xml" | awk '{print $4}' | sort -V`
    );
    const slideFiles = fileList.trim().split("\n").filter((f) => f);
    console.log(`Found ${slideFiles.length} slide files`);
    let hiddenInPresentation = [];
    try {
      const { stdout: presXml } = await execAsync$1(
        `unzip -p "${pptxPath}" "ppt/presentation.xml" 2>/dev/null`
      );
      const sldIdRegex = /<p:sldId[^>]*>/g;
      const matches = presXml.match(sldIdRegex) || [];
      matches.forEach((match, index) => {
        if (match.includes('show="0"')) {
          hiddenInPresentation.push(index + 1);
        }
      });
    } catch (e) {
      console.log("Could not parse presentation.xml for hidden slides");
    }
    for (const slideFile of slideFiles) {
      const match = slideFile.match(/slide(\d+)\.xml$/);
      if (!match) continue;
      const slideNum = parseInt(match[1], 10);
      try {
        const { stdout: slideXml } = await execAsync$1(
          `unzip -p "${pptxPath}" "${slideFile}" 2>/dev/null`
        );
        let isHidden = hiddenInPresentation.includes(slideNum);
        if (slideXml.includes('show="0"') || slideXml.includes('show="false"')) {
          isHidden = true;
        }
        const clickEffects = (slideXml.match(/nodeType="clickEffect"/g) || []).length;
        let name = `Slide ${slideNum}`;
        const titleMatch = slideXml.match(/<a:t>([^<]+)<\/a:t>/);
        if (titleMatch && titleMatch[1].length < 100) {
          name = titleMatch[1].substring(0, 50);
        }
        let notes = "";
        try {
          const slideRelsFile = slideFile.replace("slides/", "slides/_rels/").replace(".xml", ".xml.rels");
          const { stdout: relsXml } = await execAsync$1(
            `unzip -p "${pptxPath}" "${slideRelsFile}" 2>/dev/null`
          );
          const notesMatch = relsXml.match(/Type="[^"]*notesSlide"[^>]*Target="([^"]+)"/);
          if (notesMatch && notesMatch[1]) {
            const notesTarget = notesMatch[1];
            const notesFile = "ppt/" + notesTarget.replace("../", "");
            console.log(`Slide ${slideNum} notes file: ${notesFile}`);
            const { stdout: notesXml } = await execAsync$1(
              `unzip -p "${pptxPath}" "${notesFile}" 2>/dev/null`
            );
            const bodyMatch = notesXml.match(/<p:sp>.*?<p:ph type="body"[^>]*\/>.*?<p:txBody>(.*?)<\/p:txBody>.*?<\/p:sp>/s);
            if (bodyMatch && bodyMatch[1]) {
              const txBody = bodyMatch[1];
              const textMatches = txBody.match(/<a:t>([^<]*)<\/a:t>/g) || [];
              const textParts = [];
              for (const match2 of textMatches) {
                const text = match2.replace(/<\/?a:t>/g, "");
                textParts.push(text);
              }
              notes = textParts.join("").trim();
            }
            if (notes.length > 1e3) {
              notes = notes.substring(0, 1e3) + "...";
            }
            console.log(`Slide ${slideNum} notes: "${notes}"`);
          } else {
            console.log(`Slide ${slideNum}: no notes relationship found`);
          }
        } catch (e) {
          console.log(`Slide ${slideNum}: no notes file`);
        }
        const slideData = {
          slideNumber: slideNum,
          name,
          hidden: isHidden,
          animationClicks: clickEffects,
          notes
        };
        slides.push(slideData);
        if (isHidden) {
          hiddenSlides.push(slideNum);
        } else {
          visibleSlides.push(slideNum);
        }
        console.log(`Slide ${slideNum}: hidden=${isHidden}, animations=${clickEffects}`);
      } catch (e) {
        console.error(`Error parsing slide ${slideNum}:`, e);
        slides.push({
          slideNumber: slideNum,
          name: `Slide ${slideNum}`,
          hidden: false,
          animationClicks: 0,
          notes: ""
        });
        visibleSlides.push(slideNum);
      }
    }
  } catch (e) {
    console.error("Failed to parse presentation:", e);
  }
  slides.sort((a, b) => a.slideNumber - b.slideNumber);
  visibleSlides.sort((a, b) => a - b);
  hiddenSlides.sort((a, b) => a - b);
  return {
    slides,
    totalSlides: slides.length,
    visibleSlides,
    hiddenSlides
  };
}
function getNextVisibleSlide(currentSlide2, presentationData2) {
  const { visibleSlides, totalSlides: totalSlides2 } = presentationData2;
  for (const slideNum of visibleSlides) {
    if (slideNum > currentSlide2) {
      return slideNum;
    }
  }
  return null;
}
function getSlideData(slideNumber, presentationData2) {
  return presentationData2.slides.find((s) => s.slideNumber === slideNumber) || null;
}
const execAsync = util.promisify(child_process.exec);
let presentationData$1 = null;
let currentSlide$1 = 1;
let currentAnimationStep$1 = 0;
let totalSlides$1 = 1;
let currentPresentationPath = "";
let localPresentationCopy$1 = "";
function runAppleScript(script) {
  return execAsync(`osascript <<'EOF'
${script}
EOF`).then(({ stdout }) => stdout.trim());
}
function sendKeyCode(keyCode) {
  return runAppleScript(`
tell application "Microsoft PowerPoint" to activate
delay 0.1
tell application "System Events"
  key code ${keyCode}
end tell
  `);
}
async function queryCurrentSlide$1() {
  try {
    const result = await runAppleScript(`
tell application "Microsoft PowerPoint"
  try
    set ss to slide show window 1
    set currentSlideIndex to slide index of slide of slide show view of ss
    return currentSlideIndex
  on error
    return 1
  end try
end tell
    `);
    return parseInt(result.trim(), 10) || 1;
  } catch {
    return currentSlide$1;
  }
}
const macOSAutomation = {
  async checkInstalled() {
    try {
      await fs.access("/Applications/Microsoft PowerPoint.app");
      return true;
    } catch {
      return false;
    }
  },
  async openPresentation(filePath) {
    currentPresentationPath = filePath;
    const tempDir = path.join(os.tmpdir(), "slidecue-presentations");
    await fs.mkdir(tempDir, { recursive: true });
    localPresentationCopy$1 = path.join(tempDir, path.basename(filePath));
    console.log("Copying presentation to temp location...");
    await fs.copyFile(filePath, localPresentationCopy$1);
    console.log("Parsing presentation data...");
    presentationData$1 = await parsePresentationData(localPresentationCopy$1);
    console.log("Presentation data:", JSON.stringify(presentationData$1, null, 2));
    console.log("Opening presentation...");
    await execAsync(`open -a "Microsoft PowerPoint" "${localPresentationCopy$1}"`);
    await new Promise((resolve) => setTimeout(resolve, 2e3));
    try {
      const total = await runAppleScript(`
tell application "Microsoft PowerPoint"
  return count of slides of active presentation
end tell
      `);
      totalSlides$1 = parseInt(total, 10) || presentationData$1.totalSlides;
    } catch {
      totalSlides$1 = presentationData$1.totalSlides;
    }
    currentSlide$1 = 1;
    currentAnimationStep$1 = 0;
    console.log(`Opened presentation with ${totalSlides$1} slides`);
    console.log(`Visible slides: ${presentationData$1.visibleSlides.join(", ")}`);
    console.log(`Hidden slides: ${presentationData$1.hiddenSlides.join(", ")}`);
  },
  async exportThumbnails(outputDir, onProgress) {
    await fs.mkdir(outputDir, { recursive: true });
    const pptxPath = localPresentationCopy$1 || currentPresentationPath;
    const pdfPath = path.join(outputDir, "presentation.pdf");
    console.log("Converting presentation to PDF and images...");
    onProgress?.(1, totalSlides$1);
    try {
      console.log("Using LibreOffice to convert PPTX to PDF...");
      await execAsync(`/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to pdf --outdir "${outputDir}" "${pptxPath}"`);
      const baseName = path.basename(pptxPath).replace(/\.(pptx?|PPTX?)$/, "");
      const generatedPdf = path.join(outputDir, baseName + ".pdf");
      try {
        await fs.access(generatedPdf);
        if (generatedPdf !== pdfPath) {
          await execAsync(`mv "${generatedPdf}" "${pdfPath}"`);
        }
      } catch {
      }
      await fs.access(pdfPath);
      console.log("PDF created, converting to PNG...");
      await execAsync(`pdftoppm -png -r 150 "${pdfPath}" "${outputDir}/slide"`);
      const files = await fs.readdir(outputDir);
      const pngFiles = files.filter((f) => f.match(/^slide-\d+\.png$/)).sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)[0]);
        const numB = parseInt(b.match(/\d+/)[0]);
        return numA - numB;
      });
      const visibleSlides = presentationData$1?.visibleSlides || [];
      console.log("Visible slide numbers:", visibleSlides);
      console.log("PNG files to rename:", pngFiles);
      for (let i = 0; i < pngFiles.length; i++) {
        const file = pngFiles[i];
        const actualSlideNum = visibleSlides[i] || i + 1;
        const newName = `slide_${String(actualSlideNum).padStart(3, "0")}.png`;
        await execAsync(`mv "${path.join(outputDir, file)}" "${path.join(outputDir, newName)}"`);
        console.log(`Renamed ${file} -> ${newName}`);
      }
      await fs.unlink(pdfPath).catch(() => {
      });
    } catch (e) {
      console.error("Export failed:", e);
    }
    let thumbnails = [];
    try {
      const files = await fs.readdir(outputDir);
      const pngs = files.filter((f) => f.endsWith(".png")).sort();
      for (const png of pngs) {
        thumbnails.push(path.join(outputDir, png));
      }
    } catch {
    }
    console.log("Generated thumbnails:", thumbnails);
    const result = {
      thumbnails,
      totalSlides: totalSlides$1,
      hiddenSlides: presentationData$1?.hiddenSlides || [],
      visibleSlides: presentationData$1?.visibleSlides || []
    };
    return result;
  },
  async startSlideshow() {
    currentSlide$1 = 1;
    currentAnimationStep$1 = 0;
    await runAppleScript(`
tell application "Microsoft PowerPoint"
  activate
end tell
delay 0.2
tell application "System Events"
  keystroke return using {command down, shift down}
end tell
    `);
    await new Promise((resolve) => setTimeout(resolve, 500));
    currentSlide$1 = await queryCurrentSlide$1();
  },
  async nextSlide() {
    const slideData = presentationData$1 ? getSlideData(currentSlide$1, presentationData$1) : null;
    const animationsOnSlide = slideData?.animationClicks || 0;
    await sendKeyCode(124);
    await new Promise((resolve) => setTimeout(resolve, 100));
    if (currentAnimationStep$1 < animationsOnSlide) {
      currentAnimationStep$1++;
    }
    const actualSlide = await queryCurrentSlide$1();
    if (actualSlide !== currentSlide$1) {
      currentSlide$1 = actualSlide;
      currentAnimationStep$1 = 0;
      console.log(`Moved to slide ${currentSlide$1}`);
    } else {
      console.log(`Animation ${currentAnimationStep$1}/${animationsOnSlide} on slide ${currentSlide$1}`);
    }
  },
  async prevSlide() {
    await sendKeyCode(123);
    await new Promise((resolve) => setTimeout(resolve, 100));
    currentSlide$1 = await queryCurrentSlide$1();
    currentAnimationStep$1 = 0;
    console.log(`Moved to slide ${currentSlide$1}`);
  },
  async gotoSlide(slideNumber) {
    console.log(`Going to slide ${slideNumber}...`);
    try {
      if (slideNumber === 1) {
        await runAppleScript(`
tell application "Microsoft PowerPoint"
  tell slide show view of slide show window 1
    go to first slide
  end tell
end tell
        `);
      } else if (slideNumber >= totalSlides$1) {
        await runAppleScript(`
tell application "Microsoft PowerPoint"
  tell slide show view of slide show window 1
    go to last slide
  end tell
end tell
        `);
      } else {
        await runAppleScript(`
tell application "Microsoft PowerPoint"
  tell slide show view of slide show window 1
    go to first slide
    repeat until (current show position) ≥ ${slideNumber}
      go to next slide
    end repeat
  end tell
end tell
        `);
      }
    } catch (e) {
      console.error("Error in gotoSlide:", e);
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
    currentSlide$1 = await queryCurrentSlide$1();
    currentAnimationStep$1 = 0;
    console.log(`Now on slide ${currentSlide$1}`);
  },
  async getSlideInfo() {
    const actualSlide = await queryCurrentSlide$1();
    if (actualSlide !== currentSlide$1) {
      currentSlide$1 = actualSlide;
      currentAnimationStep$1 = 0;
    }
    const slideData = presentationData$1 ? getSlideData(currentSlide$1, presentationData$1) : null;
    const animationsOnSlide = slideData?.animationClicks || 0;
    const nextVisible = presentationData$1 ? getNextVisibleSlide(currentSlide$1, presentationData$1) : currentSlide$1 < totalSlides$1 ? currentSlide$1 + 1 : null;
    const currentNotes = slideData?.notes || "";
    const nextSlideData = nextVisible && presentationData$1 ? getSlideData(nextVisible, presentationData$1) : null;
    const nextNotes = nextSlideData?.notes || "";
    const visibleSlides = presentationData$1?.visibleSlides || [];
    const isLastSlide = visibleSlides.length > 0 ? currentSlide$1 === visibleSlides[visibleSlides.length - 1] : currentSlide$1 >= totalSlides$1;
    return {
      currentSlide: currentSlide$1,
      totalSlides: totalSlides$1,
      animationStep: currentAnimationStep$1,
      animationsOnSlide,
      nextVisibleSlide: nextVisible,
      isLastSlide,
      currentNotes,
      nextNotes
    };
  },
  async stopSlideshow() {
    await sendKeyCode(53);
  },
  async closePresentation() {
    await runAppleScript(`
tell application "Microsoft PowerPoint"
  close active presentation saving no
end tell
    `);
    if (localPresentationCopy$1) {
      try {
        await fs.unlink(localPresentationCopy$1);
      } catch {
      }
      localPresentationCopy$1 = "";
    }
    presentationData$1 = null;
    currentSlide$1 = 1;
    currentAnimationStep$1 = 0;
  }
};
util.promisify(child_process.exec);
let winax;
try {
  winax = require("winax");
} catch {
}
let pptApp = null;
let presentation = null;
let slideShow = null;
let presentationData = null;
let currentSlide = 1;
let currentAnimationStep = 0;
let totalSlides = 1;
let localPresentationCopy = "";
function queryCurrentSlide() {
  try {
    if (slideShow?.View) {
      return slideShow.View.CurrentShowPosition || 1;
    }
  } catch {
  }
  return currentSlide;
}
const windowsAutomation = {
  async checkInstalled() {
    try {
      const testApp = new winax.Object("PowerPoint.Application");
      testApp.Quit();
      return true;
    } catch {
      return false;
    }
  },
  async openPresentation(filePath) {
    const tempDir = path.join(os.tmpdir(), "slidecue-presentations");
    await fs.mkdir(tempDir, { recursive: true });
    localPresentationCopy = path.join(tempDir, path.basename(filePath));
    console.log("Copying presentation to temp location...");
    await fs.copyFile(filePath, localPresentationCopy);
    console.log("Parsing presentation data...");
    presentationData = await parsePresentationData(localPresentationCopy);
    console.log("Presentation data:", JSON.stringify(presentationData, null, 2));
    console.log("Opening presentation...");
    pptApp = new winax.Object("PowerPoint.Application");
    pptApp.Visible = true;
    presentation = pptApp.Presentations.Open(localPresentationCopy);
    totalSlides = presentation.Slides.Count || presentationData.totalSlides;
    currentSlide = 1;
    currentAnimationStep = 0;
    console.log(`Opened presentation with ${totalSlides} slides`);
    console.log(`Visible slides: ${presentationData.visibleSlides.join(", ")}`);
    console.log(`Hidden slides: ${presentationData.hiddenSlides.join(", ")}`);
  },
  async exportThumbnails(outputDir, onProgress) {
    await fs.mkdir(outputDir, { recursive: true });
    const paths = [];
    console.log("Exporting thumbnails...");
    onProgress?.(1, totalSlides);
    const visibleSlides = presentationData?.visibleSlides || [];
    for (let i = 0; i < visibleSlides.length; i++) {
      const slideNum = visibleSlides[i];
      const slide = presentation.Slides.Item(slideNum);
      const filePath = path.join(outputDir, `slide_${String(slideNum).padStart(3, "0")}.png`);
      try {
        slide.Export(filePath, "PNG", 1920, 1080);
        paths.push(filePath);
        console.log(`Exported slide ${slideNum}`);
      } catch (e) {
        console.error(`Failed to export slide ${slideNum}:`, e);
      }
      onProgress?.(i + 1, visibleSlides.length);
    }
    console.log("Generated thumbnails:", paths);
    return {
      thumbnails: paths,
      totalSlides,
      hiddenSlides: presentationData?.hiddenSlides || [],
      visibleSlides: presentationData?.visibleSlides || []
    };
  },
  async startSlideshow() {
    currentSlide = 1;
    currentAnimationStep = 0;
    const settings = presentation.SlideShowSettings;
    settings.StartingSlide = 1;
    settings.EndingSlide = totalSlides;
    slideShow = settings.Run();
    await new Promise((resolve) => setTimeout(resolve, 500));
    currentSlide = queryCurrentSlide();
  },
  async nextSlide() {
    const slideData = presentationData ? getSlideData(currentSlide, presentationData) : null;
    const animationsOnSlide = slideData?.animationClicks || 0;
    if (slideShow?.View) {
      slideShow.View.Next();
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
    if (currentAnimationStep < animationsOnSlide) {
      currentAnimationStep++;
    }
    const actualSlide = queryCurrentSlide();
    if (actualSlide !== currentSlide) {
      currentSlide = actualSlide;
      currentAnimationStep = 0;
      console.log(`Moved to slide ${currentSlide}`);
    } else {
      console.log(`Animation ${currentAnimationStep}/${animationsOnSlide} on slide ${currentSlide}`);
    }
  },
  async prevSlide() {
    if (slideShow?.View) {
      slideShow.View.Previous();
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
    currentSlide = queryCurrentSlide();
    currentAnimationStep = 0;
    console.log(`Moved to slide ${currentSlide}`);
  },
  async gotoSlide(slideNumber) {
    console.log(`Going to slide ${slideNumber}...`);
    if (slideShow?.View) {
      try {
        slideShow.View.GotoSlide(slideNumber);
      } catch (e) {
        console.error("GotoSlide failed, trying navigation approach:", e);
        slideShow.View.First();
        await new Promise((resolve) => setTimeout(resolve, 50));
        let pos = queryCurrentSlide();
        let iterations = 0;
        const maxIterations = totalSlides + 5;
        while (pos < slideNumber && iterations < maxIterations) {
          slideShow.View.Next();
          await new Promise((resolve) => setTimeout(resolve, 30));
          pos = queryCurrentSlide();
          iterations++;
        }
      }
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
    currentSlide = queryCurrentSlide();
    currentAnimationStep = 0;
    console.log(`Now on slide ${currentSlide}`);
  },
  async getSlideInfo() {
    const actualSlide = queryCurrentSlide();
    if (actualSlide !== currentSlide) {
      currentSlide = actualSlide;
      currentAnimationStep = 0;
    }
    const slideData = presentationData ? getSlideData(currentSlide, presentationData) : null;
    const animationsOnSlide = slideData?.animationClicks || 0;
    const nextVisible = presentationData ? getNextVisibleSlide(currentSlide, presentationData) : currentSlide < totalSlides ? currentSlide + 1 : null;
    const currentNotes = slideData?.notes || "";
    const nextSlideData = nextVisible && presentationData ? getSlideData(nextVisible, presentationData) : null;
    const nextNotes = nextSlideData?.notes || "";
    const visibleSlides = presentationData?.visibleSlides || [];
    const isLastSlide = visibleSlides.length > 0 ? currentSlide === visibleSlides[visibleSlides.length - 1] : currentSlide >= totalSlides;
    return {
      currentSlide,
      totalSlides,
      animationStep: currentAnimationStep,
      animationsOnSlide,
      nextVisibleSlide: nextVisible,
      isLastSlide,
      currentNotes,
      nextNotes
    };
  },
  async stopSlideshow() {
    if (slideShow?.View) {
      slideShow.View.Exit();
      slideShow = null;
    }
  },
  async closePresentation() {
    if (presentation) {
      try {
        presentation.Close();
      } catch {
      }
      presentation = null;
    }
    if (pptApp) {
      try {
        pptApp.Quit();
      } catch {
      }
      pptApp = null;
    }
    if (localPresentationCopy) {
      try {
        await fs.unlink(localPresentationCopy);
      } catch {
      }
      localPresentationCopy = "";
    }
    slideShow = null;
    presentationData = null;
    currentSlide = 1;
    currentAnimationStep = 0;
  }
};
function getAutomation() {
  if (process.platform === "darwin") {
    return macOSAutomation;
  } else if (process.platform === "win32") {
    return windowsAutomation;
  }
  throw new Error("Unsupported platform");
}
let server = null;
let io = null;
let currentPin = "";
let thumbnailsDir = "";
let currentPort = 3e3;
function isPortAvailable(port) {
  return new Promise((resolve) => {
    const testServer = http.createServer();
    testServer.once("error", () => resolve(false));
    testServer.once("listening", () => {
      testServer.close();
      resolve(true);
    });
    testServer.listen(port, "0.0.0.0");
  });
}
async function findAvailablePort(startPort) {
  let port = startPort;
  while (!await isPortAvailable(port)) {
    port++;
  }
  return port;
}
function getLocalIP() {
  const nets = os.networkInterfaces();
  for (const name of Object.keys(nets)) {
    for (const net of nets[name] || []) {
      if (net.family === "IPv4" && !net.internal) {
        return net.address;
      }
    }
  }
  return "localhost";
}
function generatePin() {
  return Math.floor(1e3 + Math.random() * 9e3).toString();
}
async function startServer(remoteUIPath, thumbsDir) {
  const app = express();
  server = http.createServer(app);
  io = new socket_io.Server(server, {
    cors: { origin: "*" }
  });
  currentPin = generatePin();
  thumbnailsDir = thumbsDir;
  currentPort = await findAvailablePort(3e3);
  app.use(express.static(remoteUIPath));
  app.get("/icon.png", (_req, res) => {
    const iconPath = path.join(process.cwd(), "resources", "icon.png");
    res.sendFile(iconPath);
  });
  app.use("/thumbnails", express.static(thumbnailsDir));
  app.get("/api/thumbnails", async (_req, res) => {
    try {
      const files = await fs.readdir(thumbnailsDir);
      const images = files.filter((f) => /\.(png|jpg|jpeg|gif)$/i.test(f)).sort().map((f) => `/thumbnails/${f}`);
      console.log("API thumbnails:", images);
      res.json({ thumbnails: images });
    } catch (e) {
      console.error("Error listing thumbnails:", e);
      res.json({ thumbnails: [] });
    }
  });
  io.use((socket, next) => {
    const pin = socket.handshake.auth.pin;
    if (pin === currentPin) {
      next();
    } else {
      next(new Error("Invalid PIN"));
    }
  });
  server.listen(currentPort, "0.0.0.0");
  const localIP = getLocalIP();
  return {
    url: `http://${localIP}:${currentPort}`,
    pin: currentPin,
    io
  };
}
function stopServer() {
  if (io) {
    io.close();
    io = null;
  }
  if (server) {
    server.close();
    server = null;
  }
}
function setupSocketHandlers(io2) {
  const automation2 = getAutomation();
  io2.on("connection", (socket) => {
    console.log("Remote connected:", socket.id);
    automation2.getSlideInfo().then((info) => {
      socket.emit("slide-changed", info);
    });
    socket.on("next", async () => {
      try {
        await automation2.nextSlide();
        const info = await automation2.getSlideInfo();
        io2.emit("slide-changed", info);
      } catch (error) {
        console.error("Error advancing slide:", error);
      }
    });
    socket.on("prev", async () => {
      try {
        await automation2.prevSlide();
        const info = await automation2.getSlideInfo();
        io2.emit("slide-changed", info);
      } catch (error) {
        console.error("Error going to previous slide:", error);
      }
    });
    socket.on("goto", async (slideIndex) => {
      try {
        await automation2.gotoSlide(slideIndex);
        const info = await automation2.getSlideInfo();
        io2.emit("slide-changed", info);
      } catch (error) {
        console.error("Error going to slide:", error);
      }
    });
    socket.on("disconnect", () => {
      console.log("Remote disconnected:", socket.id);
    });
  });
}
const RELEASES_REPO = {
  owner: "LAB271",
  repo: "SlideCue-releases"
};
function setupAutoUpdater(mainWindow2) {
  if (!electron.app.isPackaged) {
    console.log("Skip checkForUpdates because application is not packed and dev update config is not forced");
    return;
  }
  electronUpdater.autoUpdater.setFeedURL({
    provider: "github",
    owner: RELEASES_REPO.owner,
    repo: RELEASES_REPO.repo
  });
  electronUpdater.autoUpdater.autoDownload = false;
  electronUpdater.autoUpdater.autoInstallOnAppQuit = true;
  electronUpdater.autoUpdater.on("checking-for-update", () => {
    console.log("Checking for updates...");
  });
  electronUpdater.autoUpdater.on("update-available", async (info) => {
    console.log("Update available:", info.version);
    const result = await electron.dialog.showMessageBox(mainWindow2, {
      type: "info",
      title: "Update Available",
      message: `A new version (${info.version}) is available!`,
      detail: "Would you like to download and install it now?",
      buttons: ["Download", "Later"],
      defaultId: 0,
      cancelId: 1
    });
    if (result.response === 0) {
      electronUpdater.autoUpdater.downloadUpdate();
    }
  });
  electronUpdater.autoUpdater.on("update-not-available", () => {
    console.log("No updates available - you have the latest version");
  });
  electronUpdater.autoUpdater.on("download-progress", (progress) => {
    console.log(`Download progress: ${Math.round(progress.percent)}%`);
    mainWindow2.setProgressBar(progress.percent / 100);
  });
  electronUpdater.autoUpdater.on("update-downloaded", async (info) => {
    console.log("Update downloaded:", info.version);
    mainWindow2.setProgressBar(-1);
    const result = await electron.dialog.showMessageBox(mainWindow2, {
      type: "info",
      title: "Update Ready",
      message: "Update downloaded!",
      detail: "The update will be installed when you restart SlideCue.",
      buttons: ["Restart Now", "Later"],
      defaultId: 0,
      cancelId: 1
    });
    if (result.response === 0) {
      electronUpdater.autoUpdater.quitAndInstall();
    }
  });
  electronUpdater.autoUpdater.on("error", (error) => {
    console.error("Auto-updater error:", error.message);
  });
  setTimeout(() => {
    electronUpdater.autoUpdater.checkForUpdates().catch((err) => {
      console.error("Failed to check for updates:", err.message);
    });
  }, 3e3);
}
let mainWindow = null;
let presentationFile = null;
let thumbnailPaths = [];
let currentThumbsDir = "";
const automation = getAutomation();
function createWindow() {
  mainWindow = new electron.BrowserWindow({
    width: 1200,
    height: 800,
    fullscreen: true,
    webPreferences: {
      preload: path.join(__dirname, "../preload/index.js"),
      contextIsolation: true,
      nodeIntegration: false,
      webSecurity: false
      // Allow loading local files
    }
  });
  if (process.env.NODE_ENV === "development") {
    mainWindow.loadURL("http://localhost:5173");
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, "../renderer/index.html"));
  }
  return mainWindow;
}
electron.protocol.registerSchemesAsPrivileged([
  { scheme: "slidecue", privileges: { secure: true, supportFetchAPI: true, stream: true } }
]);
electron.app.whenReady().then(async () => {
  electron.protocol.handle("slidecue", (request) => {
    const filePath = request.url.replace("slidecue://", "");
    return electron.net.fetch("file://" + filePath);
  });
  const window = createWindow();
  const isPPTInstalled = await automation.checkInstalled();
  if (!isPPTInstalled) {
    electron.dialog.showErrorBox(
      "PowerPoint Not Found",
      "Microsoft PowerPoint is required to run presentations. Please install PowerPoint and restart SlideCue."
    );
  }
  if (process.env.NODE_ENV !== "development") {
    setupAutoUpdater(window);
  }
  electron.app.on("activate", () => {
    if (electron.BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});
electron.ipcMain.handle("check-powerpoint", async () => {
  return automation.checkInstalled();
});
electron.ipcMain.handle("import-presentation", async () => {
  const result = await electron.dialog.showOpenDialog(mainWindow, {
    filters: [{ name: "PowerPoint", extensions: ["pptx", "ppt"] }],
    properties: ["openFile"]
  });
  if (result.canceled || !result.filePaths[0]) {
    return null;
  }
  presentationFile = result.filePaths[0];
  currentThumbsDir = path.join(
    os.tmpdir(),
    "slidecue-thumbs",
    Date.now().toString()
  );
  await fs.mkdir(currentThumbsDir, { recursive: true });
  const sendProgress = (step, total, message) => {
    mainWindow?.webContents.send("import-progress", { step, total, message });
  };
  sendProgress(0, 5, "Starting...");
  sendProgress(1, 5, "Opening presentation...");
  await automation.openPresentation(presentationFile);
  sendProgress(2, 5, "Getting slide count...");
  await automation.getSlideInfo();
  sendProgress(3, 5, "Converting to PDF...");
  const slideMetadata = await automation.exportThumbnails(currentThumbsDir, (current, total) => {
    sendProgress(3 + current / total, 5, `Converting slide ${current}/${total}...`);
  });
  thumbnailPaths = slideMetadata.thumbnails;
  sendProgress(5, 5, "Done!");
  console.log("Thumbnails exported to:", currentThumbsDir);
  console.log("Thumbnail files:", thumbnailPaths);
  console.log("Hidden slides:", slideMetadata.hiddenSlides);
  return {
    filePath: presentationFile,
    fileName: path.basename(presentationFile),
    thumbnails: thumbnailPaths,
    totalSlides: slideMetadata.totalSlides,
    visibleSlides: slideMetadata.visibleSlides,
    hiddenSlides: slideMetadata.hiddenSlides
  };
});
electron.ipcMain.handle("start-presentation", async () => {
  if (!presentationFile) {
    throw new Error("No presentation loaded");
  }
  const remoteUIPath = electron.app.isPackaged ? path.join(process.resourcesPath, "remote") : path.join(electron.app.getAppPath(), "resources/remote");
  console.log("Remote UI path:", remoteUIPath);
  console.log("Thumbnails dir:", currentThumbsDir);
  const { url, pin, io: io2 } = await startServer(remoteUIPath, currentThumbsDir);
  setupSocketHandlers(io2);
  await automation.startSlideshow();
  return {
    url,
    pin,
    localIP: getLocalIP()
  };
});
electron.ipcMain.handle("stop-presentation", async () => {
  await automation.stopSlideshow();
  stopServer();
});
electron.ipcMain.handle("get-slide-info", async () => {
  return automation.getSlideInfo();
});
function cleanupTempDirs() {
  const tempBase = os.tmpdir();
  const dirsToClean = ["slidecue-thumbs", "slidecue-presentations"];
  for (const dir of dirsToClean) {
    const dirPath = path.join(tempBase, dir);
    try {
      fsSync.rmSync(dirPath, { recursive: true, force: true });
      console.log(`Cleaned up: ${dirPath}`);
    } catch (e) {
    }
  }
}
electron.app.on("will-quit", () => {
  cleanupTempDirs();
});
electron.app.on("window-all-closed", () => {
  stopServer();
  electron.app.quit();
});
