const puppeteer = require("puppeteer");
const express = require("express");
const clc = require("cli-color");
const ExcelJS = require("exceljs");
const cookieParser = require("cookie-parser");
const firstname_input = "Nathan";
const lastname_input = "Matounga";
const pdf_input = "nathan.pdf";
const email_input = "nathan.matounga@epitech.eu";
const app = express();

(async () => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  app.use(cookieParser());

  await page.goto(
    "https://www.hellowork.com/fr-fr/emploi/recherche.html?k=D%C3%A9veloppeur&p=8&mode=pagination"
  );

  await page.waitForSelector("#hw-cc-notice-accept-btn", { visible: true });

  await page.click("#hw-cc-notice-accept-btn");

  const workbooks = new ExcelJS.Workbook();
  await workbooks.xlsx.readFile("data.xlsx");
  const worksheets = workbooks.getWorksheet("Data");
  const jobColumn = worksheets.getColumn(3);
  const titleColumn = worksheets.getColumn(1);
  const anonceColumn = worksheets.getColumn(2);

  const city_array = [];
  for (const row of jobColumn._worksheet._rows) {
    city_array.push(row.getCell(jobColumn._number).value);
  }

  const job_array = [];
  for (const row of titleColumn._worksheet._rows) {
    job_array.push(row.getCell(titleColumn._number).value);
  }

  const anonce_array = [];
  for (const row of anonceColumn._worksheet._rows) {
    anonce_array.push(row.getCell(anonceColumn._number).value);
  }

  const links = await page.$$eval(
    "a",
    (elements, keyword1, keyword2, keyword3, keyword4, keyword5) => {
      return elements
        .filter(
          (element) =>
            element.textContent.includes(keyword1) ||
            element.textContent.includes(keyword2) ||
            element.textContent.includes(keyword3) ||
            element.textContent.includes(keyword4) ||
            element.textContent.includes(keyword5)
        )
        .map((element) => element.href);
    },
    "Developpeur",
    "developpeur",
    "web",
    "développeur",
    "Développeur",
    "Développeur.e",
    "Développeuse",
    "développeur.e",
    "developpeur.e",
    "Developpeur.e",
    "Programmeur",
    "programmeur",
    "php",
    "PHP",
    "JS",
    "js",
    "javascript"
  );

  for (const link of links) {
    await page.goto(link);

    const postulerButtons = await page.$$eval(
      "a[href*='#postuler']",
      (elements) => {
        return elements.map((element) => element.href);
      }
    );
    if (postulerButtons.length > 0) {
      const firstPostulerButton = postulerButtons[0];
      await page.goto(firstPostulerButton);

      const job = await page.$eval(
        "body > main > div.tw-relative > div.tw-layout-grid.tw-mt-6.sm\\:tw-mt-12 > div > div.tw-col-span-full.lg\\:tw-col-span-8 > div > div > h1 > span.tw-contents.tw-typo-m.tw-text-grey",
        (element) => element.textContent
      );
      const city = await page.$eval(
        "body > main > div.tw-relative > div.tw-layout-grid.tw-mt-6.sm\\:tw-mt-12 > div > div.tw-col-span-full.lg\\:tw-col-span-8 > div > div > span:nth-child(3)",
        (element) => element.textContent
      );
      const title = await page.$eval(
        "body > main > div.tw-relative > div.tw-layout-grid.tw-mt-6.sm\\:tw-mt-12 > div > div.tw-col-span-full.lg\\:tw-col-span-8 > div > div > h1 > span.tw-block.tw-typo-xl.sm\\:tw-typo-3xl.tw-mb-2",
        (element) => element.textContent
      );

      // Middleware pour supprimer tous les cookies
      app.use(function (req, res) {
        // Récupérer la liste des cookies de la requête
        var cookies = req.cookies;
        // Parcourir chaque cookie et le supprimer
        for (var cookie in cookies) {
          res.clearCookie(cookie);
        }
      });

      const currentDate = new Date();
      const addLeadingZero = (number) => {
        return number < 10 ? "0" + number : number;
      };

      try {
        let city_verif = false;
        let title_verif = false;
        let anonce_verif = false;

        const city_filter = city_array.includes(city.trim());

        const title_filter = job_array.includes(job.trim());
        const anonce_filter = anonce_array.includes(title.trim());

        if (city_filter && title_filter && anonce_array) {
          console.log("Annonce déja faite je passe à la suivante");
          continue;
        }
      } catch (error) {
        console.log("Erreur :", error.message);
      }

      const day = addLeadingZero(currentDate.getDate());
      const month = addLeadingZero(currentDate.getMonth() + 1);
      const year = currentDate.getFullYear();
      const hours = addLeadingZero(currentDate.getHours());
      const minutes = addLeadingZero(currentDate.getMinutes());
      const seconds = addLeadingZero(currentDate.getSeconds());
      const formattedDate = `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;

      console.log(clc.yellow("Entreprise: " + job));
      console.log(clc.yellow("Lieu : " + city));
      console.log(clc.yellow("Titre annonce : " + title));
      console.log(
        clc.yellow("Date d'envoie de la candidature" + formattedDate)
      );
      console.log(clc.yellow("Remplissage du formulaire"));

      await Promise.all([
        page.waitForSelector('input[name="Firstname"]', { timeout: 5000 }),
        page.waitForSelector('input[name="Name"]', { timeout: 5000 }),
        page.waitForSelector('input[name="Email"]', { timeout: 5000 }),
        page.waitForSelector('input[type="button"]', { timeout: 5000 }),
        page.waitForSelector('button[type="submit"]', { timeout: 5000 }),
      ]);

      try {
        await page.type('input[name="Firstname"]', firstname_input, {
          delay: 200,
        });
        console.log(clc.green("Remplissage du champs Firstname"));
      } catch (error) {
        console.log(clc.red("Erreur de remplissage du champs Firstname"));
      }
      try {
        await page.type('input[name="Name"]', lastname_input, { delay: 200 });
        console.log(clc.green("Remplissage du champs Name"));
      } catch (error) {
        console.log(clc.red("Erreur de remplissage du champs Lastname"));
      }
      try {
        await page.type('input[name="Email"]', email_input, {
          delay: 200,
        });
        console.log(clc.green("Remplissage du champs Email"));
      } catch (error) {
        console.log(clc.red("Erreur de remplissage du champs Email"));
      }
      try {
        const fileInput = await page.$("#rjupload");
        const filePath = pdf_input;
        await fileInput.uploadFile(filePath);
        console.log(clc.green("Remplissage du champs CV"));
      } catch (error) {
        console.log(clc.red("Remplissage du champs CV impossible."));
      }
      await new Promise((resolve) => setTimeout(resolve, 2000));

      try {
        await page.waitForSelector('button[data-cy="submitButton"]', {
          visible: true,
          timeout: 5000,
        });

        await page.click('button[data-cy="submitButton"]');
        console.log(clc.green("Le bouton a été cliqué avec succès."));
      } catch (error) {
        console.log(
          clc.redBright(
            "Je ne peux pas cliquer sur le bouton ou une autre erreur s'est produite."
          )
        );
      }
      await new Promise((resolve) => setTimeout(resolve, 3000));

      try {
        const currentURL = page.url();
        console.log("URL actuelle de la page après redirection:", currentURL);
      } catch (error) {
        console.log("url introuvable");
      }

      const workbook = new ExcelJS.Workbook();
      workbook.xlsx
        .readFile("data.xlsx")
        .then(() => {
          const worksheet = workbook.getWorksheet("Data");
          worksheet.columns = [
            { header: "Entreprise", key: "job" },
            { header: "Titre de l'annonce", key: "title" },
            { header: "Ville", key: "city" },
            { header: "Date", key: "date" },
          ];

          const data = [
            { job: job, title: title, city: city, date: formattedDate },
          ];

          data.forEach((row) => {
            worksheet.addRow(row);
          });
          return workbook.xlsx.writeFile("data.xlsx");
        })
        .then(() => {
          console.log(
            "Les données ont été ajoutées au fichier Excel avec succès."
          );
        })
        .catch((error) => {
          console.log(
            "Une erreur s'est produite lors de l'ajout des données au fichier Excel :",
            error
          );
        });
      console.log(clc.bgCyan("\t\t\tJe passe au liens suivant"));
    } else {
      console.log(clc.red("Aucun bouton de postulation trouvé sur cette page"));
      console.log(clc.bgCyan("\t\t\tJe passe au liens suivant"));
    }
  }
  await browser.close();
})();
