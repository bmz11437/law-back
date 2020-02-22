const Koa = require("koa");
const koaBody = require("koa-body");
const path = require("path");
const Router = require("koa-router");
const app = new Koa();
let router = new Router();
let {  getQszDocs } = require("./exportDoc");
router.post("/word", async ctx => {
  let postParam = ctx.request.body;
  await getQszDocs(postParam);
  ctx.body = "导出成功";
});

app.use(
  koaBody({
    multipart: true,
    formidable: {
      uploadDir: path.join(__dirname, "/public/uploads"),
      keepExtensions: true
    }
  })
);

app.use(router.routes());
app.listen(3000, () => console.log("程序启动在 3000 端口了"));
