const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

(async () => {
  const websiteLink = "WEBSITE_LINK_HERE";

  const browser = await puppeteer.launch({
    headless: false,
    slowMo: 100,
  });
  const page = await browser.newPage();

  await page.goto(websiteLink, {
    waitUntil: "domcontentloaded",
  });

  let allUrls = [];

  await page.waitForSelector(".product.product--teaser");

  const productUrls = await page.evaluate(() => {
    return Array.from(
      document.querySelectorAll(".product.product--teaser")
    ).map((el) => el.getAttribute("data-product-url"));
  });

  allUrls.push(...productUrls);

  let productsData = [];

  for (const url of allUrls) {
    const productPage = await browser.newPage();
    await productPage.goto(url, {
      waitUntil: "domcontentloaded",
    });

    // Extract product details
    const productData = await productPage.evaluate(() => {
      const title =
        document.querySelector(".product__title")?.innerText || "N/A";
      const price =
        document.querySelector(".product__price-value")?.innerText || "N/A";

      const sizes = Array.from(
        document.querySelectorAll(".product__option-values label")
      ).map((el) => el.innerText.trim());

      const descriptions = document.querySelectorAll(".description");

      let composition = "";
      let descriptionText = "";

      descriptions.forEach((description) => {
        const compositionParagraph = Array.from(
          description.querySelectorAll("p")
        ).find((p) => p.textContent.includes("Composition :"));

        if (compositionParagraph) {
          const compositionText = compositionParagraph.textContent
            .replace("Composition : ", "")
            .trim();
          composition = compositionText;

          const elementsBeforeComposition = Array.from(
            description.children
          ).slice(
            0,
            Array.from(description.children).indexOf(compositionParagraph)
          );

          const textBeforeComposition = elementsBeforeComposition
            .map((element) => {
              if (element.tagName === "P") {
                return element.textContent.trim();
              } else if (element.tagName === "UL") {
                return Array.from(element.querySelectorAll("li"))
                  .map((li) => li.textContent.trim())
                  .join(" ");
              }
            })
            .join(" ");

          descriptionText = textBeforeComposition;
        }
      });

      return { title, price, sizes, descriptionText, composition };
    });

    productsData.push({ url, ...productData });

    await productPage.close();
  }

  //   Write data to an Excel file
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Products");

  // Define columns for the Excel file
  worksheet.columns = [
    { header: "Title", key: "title", width: 40 },
    { header: "Price", key: "price", width: 20 },
    { header: "Sizes", key: "sizes", width: 30 },
    { header: "Description", key: "descriptionText", width: 80 },
    { header: "Composition", key: "composition", width: 30 },
    { header: "URL", key: "url", width: 100 },
  ];

  // Add data to the worksheet
  productsData.forEach((item) => {
    worksheet.addRow({
      title: item.title,
      price: item.price,
      sizes: item.sizes.join(", "), // Join sizes into a comma-separated string
      descriptionText: item.descriptionText,
      composition: item.composition,
    });
  });

  // Save the Excel file
  await workbook.xlsx.writeFile("products.xlsx");
  console.log("Excel file has been generated!");
})();
