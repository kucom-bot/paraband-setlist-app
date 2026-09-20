// ไฟล์: api/gemini.js
export default async function handler(req, res) {
  // รับเฉพาะ HTTP POST เท่านั้น
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method Not Allowed' });
  }

  try {
    // 1. รับข้อมูลรูปภาพจากหน้าเว็บเรา
    const { base64Data, mimeType } = req.body;
    if (!base64Data || !mimeType) {
      return res.status(400).json({ success: false, message: 'Missing image data' });
    }

    // 2. ดึง API Key จากที่ซ่อนไว้ (Vercel จะรู้จัก process.env อัตโนมัติ)
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ success: false, message: 'API Key not configured on server' });
    }

    // 3. ยิง API ไปหา Gemini จากหลังบ้าน
    const response = await fetch(`https://generativelanguage.googleapis.com/v1/models/gemini-3.6-flash:generateContent?key=${apiKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{
          parts: [
            { text: "Extract the lyrics and chords from this image. Output ONLY the raw text. Do NOT use markdown code blocks. Preserve the exact placement of chords relative to the lyrics." },
            { inline_data: { mime_type: mimeType, data: base64Data } }
          ]
        }]
      })
    });

    const data = await response.json();
    if (data.error) throw new Error(data.error.message);

    // 4. ส่งข้อความที่แปลงเสร็จกลับไปให้หน้าเว็บเรา
    let extractedText = data.candidates[0].content.parts[0].text;
    res.status(200).json({ success: true, text: extractedText });

  } catch (error) {
    console.error("Vercel Function Error:", error);
    res.status(500).json({ success: false, message: error.message });
  }
}