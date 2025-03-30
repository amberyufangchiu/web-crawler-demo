const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

(async () => {
  const websiteLink = "WEBSITE_LINK_HERE";
  const browser = await puppeteer.launch({
    headless: true,
    slowMo: 100,
    ignoreHTTPSErrors: true,
    defaultViewport: null,
    ignoreDefaultArgs: ["--enable-automation"],
    args: [
      "--disable-infobars",
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
      "--disable-gpu",
      "--enable-webgl",
      "--window-size=1600,900",
      "--start-maximized",
      // Args for AWS Lambda:
      "--single-process",
      "--user-data-dir=/tmp/user-data",
      "--data-path=/tmp/data-path",
      "--homedir=/tmp",
      "--disk-cache-dir=/tmp/cache-dir",
      "--database=/tmp/database",
    ],
  });
  const page = await browser.newPage();

  await page.goto(websiteLink, {
    waitUntil: "domcontentloaded",
    timeout: 120000,
  });

  await page.waitForSelector(".pagination");

  const totalPages = await page.evaluate(() => {
    const pages = document
      .querySelector(".pagination li:nth-last-child(2) a")
      ?.getAttribute("data-page");

    return parseInt(pages);
  });

  await page.close();

  console.log(`Total Pages Found: ${totalPages}`);

  // Set up Excel workbook
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Products");

  worksheet.columns = [
    { header: "Title", key: "title", width: 40 },
    { header: "Price", key: "price", width: 20 },
    { header: "Sizes", key: "sizes", width: 30 },
    { header: "Description", key: "descriptionText", width: 80 },
    { header: "Composition", key: "composition", width: 30 },
    { header: "URL", key: "url", width: 100 },
  ];

  // Fetch product URLs
  for (let i = 1; i <= totalPages; i++) {
    const categoryPage = await browser.newPage();
    console.log(`Scraping page ${i} of ${totalPages}...`);

    try {
      await categoryPage.setUserAgent(
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36"
      );
      await categoryPage.goto(`${websiteLink}?page=${i}`, {
        waitUntil: "networkidle2",
        timeout: 120000,
      });
      await categoryPage.waitForSelector(".product.product--teaser", {
        timeout: 60000,
      });
    } catch (error) {
      console.error(`Error loading category page ${i}:`, error);
      await categoryPage.close();
      continue;
    }

    const productUrls = await categoryPage.evaluate(() => {
      return Array.from(
        document.querySelectorAll(".product.product--teaser")
      ).map((el) => el.getAttribute("data-product-url"));
    });

    console.log(`page ${i} all products: ${productUrls}`);

    let products = [];

    for (const url of productUrls) {
      console.log(`Scraping product: ${url}`);

      const productPage = await browser.newPage();
      let retries = 3;
      while (retries > 0) {
        try {
          await productPage.setUserAgent(
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36"
          );

          await productPage.goto(url, {
            waitUntil: "networkidle2",
            timeout: 180000,
          });

          const productData = await productPage.evaluate(() => {
            const title =
              document.querySelector(".product__title")?.innerText || "N/A";
            const price =
              document.querySelector(".product__price-value")?.innerText ||
              "N/A";
            const sizes = Array.from(
              document.querySelectorAll(".product__option-values label")
            ).map((el) => el.innerText.trim());

            let descriptionText = "";
            let composition = "";
            const descriptions = document.querySelectorAll(".description");
            descriptions.forEach((description) => {
              const compositionParagraph = Array.from(
                description.querySelectorAll("p")
              ).find((p) => p.textContent.includes("Composition :"));
              if (compositionParagraph) {
                composition = compositionParagraph.textContent
                  .replace("Composition : ", "")
                  .trim();
                const elementsBeforeComposition = Array.from(
                  description.children
                ).slice(
                  0,
                  Array.from(description.children).indexOf(compositionParagraph)
                );
                descriptionText = elementsBeforeComposition
                  .map((element) =>
                    element.tagName === "P" ? element.textContent.trim() : ""
                  )
                  .join(" ");
              }
            });
            return { title, price, sizes, descriptionText, composition };
          });

          products.push({ url, ...productData });
          break;
        } catch (error) {
          console.error(
            `Error scraping product ${url} (Attempts left: ${retries - 1}):`,
            error
          );
          retries--;
          if (retries === 0) {
            console.error(
              `Skipping product ${url} after multiple failed attempts.`
            );
          }
        }
      }
      await productPage.close();
    }

    console.log(`Writing ${products.length} products to Excel...`);

    products.forEach((item) => {
      worksheet.addRow({
        title: item.title,
        price: item.price,
        sizes: item.sizes.join(", "),
        descriptionText: item.descriptionText,
        composition: item.composition,
        url: item.url,
      });
    });

    // Clear memory after writing
    products = [];

    await workbook.xlsx.writeFile("products.xlsx");
    console.log("Excel file updated!");

    await categoryPage.close();
  }

  console.log("Scraping completed. Final Excel file saved!");

  await browser.close();
})();
