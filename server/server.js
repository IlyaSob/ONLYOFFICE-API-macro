//import all the packages
import express from "express";
import SerpApi from "google-search-results-nodejs";
import cors from "cors";
import oData from "./sampleData.js";
import dotenv from "dotenv";
//cors setup to fix all CORS related errors on the client side
const app = express();
app.use(cors());

//dotenv config spec
dotenv.config();
const search = new SerpApi.GoogleSearch(process.env.SERP_API_KEY);

// GET route for the search
app.get("/search", async (req, res, next) => {
  const { query } = req.query;
  console.log(query);
  try {
    const params = {
      engine: "baidu",
      q: query,
    };

    search.json(params, function (data) {
      console.log(data);
      res.send(data);
    });
  } catch (error) {
    console.log(error);
    res
      .status(500)
      .json({ error: "An error has occured. Check log for more details." });
  }
});

app.get("/test", async (req, res, next) => {
  const { query } = req.query;
  try {
    res.json(oData);
  } catch (error) {
    console.log(error);
    res
      .status(500)
      .json({ error: "An error has occured. Check log for more details." });
  }
});

app.get("/advancedSearch", async (req, res, next) => {
  const { query, rn, pn } = req.query;
  let { no_cache } = req.query;
  console.log(query, rn, pn, no_cache);
  if (no_cache === "Y" || no_cache === "y") {
    no_cache = true;
  } else if (no_cache === "N" || no_cache === "n") {
    no_cache = false;
  } else {
    no_cache = false;
  }
  try {
    const params = {
      engine: "baidu",
      q: query,
      rn: rn,
      pn: pn,
      no_cache: no_cache,
    };

    search.json(params, function (data) {
      console.log(data);
      res.send(data);
    });
  } catch (error) {
    console.log(error);
    res
      .status(500)
      .json({ error: "An error has occured. Check log for more details." });
  }
});

app.listen(3000, () => {
  console.log("Server now setup on PORT 3000");
});
