// ไฟล์: api/gemini.js
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method Not Allowed' });
  }

  try {
    const { base64Data, mimeType } = req.body;
    if (!base64Data || !mimeType) return res.status(400).json({ success: false, message: 'Missing image data' });

    // 1. ดึงคีย์ทั้งหมดที่คั่นด้วยลูกน้ำมาจาก Vercel
    const keysString = process.env.GEMINI_API_KEYS;
    if (!keysString) return res.status(500).json({ success: false, message: 'API Keys not configured' });

    // 2. หั่นลูกน้ำแยกเป็นกล่องๆ (Array)
    const keysArray = keysString.split(',');
    
    // 3. สุ่มเลือกคีย์มา 1 อันเพื่อใช้งานในรอบนี้ (ไม้ตายลดอัตราการติด Limit)
    const randomKey = keysArray[Math.floor(Math.random() * keysArray.length)].trim();

    const response = await fetch(`https://generativelanguage.googleapis.com/v1/models/gemini-3.6-flash:generateContent?key=${randomKey}`, {
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

    let extractedText = data.candidates[0].content.parts[0].text;
    res.status(200).json({ success: true, text: extractedText });

  } catch (error) {
    console.error("Vercel Function Error:", error);
    res.status(500).json({ success: false, message: error.message });
  }
}