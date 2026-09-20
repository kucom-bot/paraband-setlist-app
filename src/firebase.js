import { initializeApp } from "firebase/app";
import { getFirestore, enableIndexedDbPersistence } from "firebase/firestore";

// แทนที่ตรงนี้ด้วย Config ของคุณที่ได้จาก Firebase
const firebaseConfig = {
  apiKey: "AIzaSyBz2VLL2F6qzVXt07G-mq_BJEJ1RiIL1Wc",
  authDomain: "bandsetlistapp.firebaseapp.com",
  projectId: "bandsetlistapp",
  storageBucket: "bandsetlistapp.firebasestorage.app",
  messagingSenderId: "428907797276",
  appId: "1:428907797276:web:e122c9fbd5ea7aa799f8d9",
  measurementId: "G-DG1Q9FEBK0"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// เปิดใช้งาน Offline Persistence (จุดสำคัญสำหรับ Offline Mode เวลางานเล่นสดเน็ตหลุด)
enableIndexedDbPersistence(db)
  .catch((err) => {
    if (err.code == 'failed-precondition') {
      console.log('Multiple tabs open, persistence can only be enabled in one tab at a a time.');
    } else if (err.code == 'unimplemented') {
      console.log('The current browser does not support all of the features required to enable persistence');
    }
  });

export { db };