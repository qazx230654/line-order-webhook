const express = require("express");

const app = express();

app.use(express.json());

app.post("/webhook", (req, res) => {

  console.log(req.body);

  res.status(200).send("OK");

});

app.get("/", (req, res) => {
  res.send("OK");
});

const port =
  process.env.PORT || 8080;

app.listen(port);