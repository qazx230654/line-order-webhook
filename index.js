const functions =
  require('@google-cloud/functions-framework');

const crypto =
  require('crypto');

const axios =
  require('axios');

function getRequiredEnv(name){
  const value = process.env[name];

  if(!value){
    throw new Error(`Missing environment variable: ${name}`);
  }

  return value;
}

const GAS_URL =
  getRequiredEnv('GAS_URL');

const GAS_SHARED_SECRET =
  getRequiredEnv('GAS_SHARED_SECRET');

const CLOUD_RUN_NOTIFY_SECRET =
  getRequiredEnv('CLOUD_RUN_NOTIFY_SECRET');

const LINE_CHANNEL_ACCESS_TOKEN =
  getRequiredEnv('LINE_CHANNEL_ACCESS_TOKEN');

const LINE_CHANNEL_SECRET =
  getRequiredEnv('LINE_CHANNEL_SECRET');

const LINE_API_TIMEOUT_MS = 10000;
const GAS_TIMEOUT_MS = 45000;

const lineApi =
  axios.create({
    baseURL:'https://api.line.me/v2/bot',
    timeout:LINE_API_TIMEOUT_MS,
    headers:{
      Authorization:
        `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`
    }
  });

const gasApi =
  axios.create({
    timeout:GAS_TIMEOUT_MS
  });

function verifyLineSignature(req){

  const signature =
    req.get
      ? req.get('x-line-signature')
      : req.headers['x-line-signature'];

  if(
    !signature ||
    !Buffer.isBuffer(req.rawBody)
  ){
    return false;
  }

  const expected =
    crypto
      .createHmac(
        'sha256',
        LINE_CHANNEL_SECRET
      )
      .update(req.rawBody)
      .digest('base64');

  const actualBuffer =
    Buffer.from(signature);

  const expectedBuffer =
    Buffer.from(expected);

  return (
    actualBuffer.length ===
      expectedBuffer.length &&
    crypto.timingSafeEqual(
      actualBuffer,
      expectedBuffer
    )
  );

}

functions.http(
  'helloHttp',
  async (req, res) => {

    try{

      // 完成通知
      if(req.path === "/notifyReady"){

        try{

          if(
            req.method !== "POST" ||
            !req.body ||
            req.body.notifySecret !==
              CLOUD_RUN_NOTIFY_SECRET
          ){
            return res
              .status(401)
              .send("UNAUTHORIZED");
          }

          const userId =
            req.body.userId;

          const orderId =
            req.body.orderId;

          await lineApi.post(
            '/message/push',
            {
              to:userId,
              messages:[
                {
                  type:"text",
                  text:
`🎉 您的訂單已完成

訂單編號：${orderId}

歡迎前來取餐！`
                }
              ]
            }
          );

          return res
            .status(200)
            .send("OK");

        }catch(error){

          console.error(error);

          return res
            .status(500)
            .send("ERROR");

        }

      }

      if(req.method !== "POST"){
        return res
          .status(200)
          .send("OK");
      }

      if(!verifyLineSignature(req)){
        return res
          .status(401)
          .send("INVALID_SIGNATURE");
      }

      const events =
        req.body.events || [];

      for(const event of events){

        if(
          event.type !== "message"
        )
          continue;

        if(
          event.message.type !== "text"
        )
          continue;

        if(
          event.source.type === "group"
        )
          continue;

        const text =
          event.message.text.trim();

        // 忽略聊天訊息
        const ignoreWords = [

          "好",
          "收到",
          "謝謝",
          "OK",
          "Ok",
          "ok",
          "👌",
          "👍"

        ];

        if(
          ignoreWords.includes(text)
        ){
          continue;
        }

        const userId =
          event.source.userId;

        const profile =
          await lineApi.get(
            `/profile/${userId}`
          );

        const displayName =
          profile.data.displayName;

        const pictureUrl =
          profile.data.pictureUrl || "";

        console.log(
          displayName,
          text
        );

        const response =
          await gasApi.post(
            GAS_URL,
            {

              message:text,

              customerName:
                displayName,

              lineUserId:
                userId,

              pictureUrl:
                pictureUrl,

              webhookSecret:
                GAS_SHARED_SECRET,

              webhookEventId:
                event.webhookEventId || ""

            }
          );

        const result =
          response.data;

        console.log(result);

        if(result.error === "UNAUTHORIZED"){
          throw new Error(
            "Apps Script webhook authentication failed"
          );
        }

        if(result.duplicate){

          await lineApi.post(
            '/message/reply',
            {
              replyToken:event.replyToken,
              messages:[
                {
                  type:"text",
                  text:
                    "這筆訊息已經處理過，不會重複建立訂單。"
                }
              ]
            }
          );

          continue;

        }

        if(
          !result.is_order
        ){

          continue;

        }

        if(result.soldOut){

          await lineApi.post(
            '/message/reply',
            {
              replyToken:
                event.replyToken,

              messages:[
                {
                  type:"text",
                  text:
        `不好意思，以下品項今日已售完：\n\n${result.soldOutItems.join("、")}\n\n請重新傳送完整訂單。`
                }
              ]
            }
          );

          continue;

        }

        let orderText = "";

        result.order.groups.forEach(group=>{

          orderText +=
            `【${group.name || "一般"}】\n`;

          group.items.forEach(item=>{

            orderText +=
              `${item.name} x${item.quantity}\n`;

          });

          if(group.note){

            orderText +=
              `備註：${group.note}\n`;

          }

          orderText += "\n";

        });

        let priceText = "";

        if(
          result.priceMissing
        ){

          priceText =
`⚠️ 部分品項尚未設定價格

${result.missingItems.join("、")}`;

        }else{

          priceText =
`💰 預估金額：${result.totalPrice} 元`;

        }

        await lineApi.post(

          '/message/reply',

          {

            replyToken:
              event.replyToken,

            messages:[
              {

                type:"text",

                text:
`✅ 已收到您的訂單

${orderText}
${priceText}

🕒 預計取餐：
${result.estimatedPickupTime}

訂單編號：
${result.orderId}

如需修改，
請在 20 分鐘內重新傳送完整訂單。`

              }

            ]

          }

        );

      }

      res
        .status(200)
        .send("OK");

    }catch(error){

      console.error(error);

      if(error.response){

        console.error(
          error.response.data
        );

      }

      res
        .status(500)
        .send("ERROR");

    }

  }

);
