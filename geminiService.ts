
import { GoogleGenAI, Type, GenerateContentResponse } from "@google/genai";
import { LessonData } from './types';

const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

const lessonSchema = {
  type: Type.OBJECT,
  properties: {
    slides: {
      type: Type.ARRAY,
      minItems: 10,
      maxItems: 10,
      items: {
        type: Type.OBJECT,
        properties: {
          title: { type: Type.STRING },
          content: { type: Type.STRING },
          imagePrompt: { type: Type.STRING, description: "Detailed visual prompt for generating an educational image for this slide. Focus on subjects and clarity." }
        },
        required: ["title", "content", "imagePrompt"]
      }
    },
    tests: {
      type: Type.ARRAY,
      minItems: 10,
      maxItems: 10,
      items: {
        type: Type.OBJECT,
        properties: {
          question: { type: Type.STRING },
          options: { type: Type.ARRAY, items: { type: Type.STRING }, minItems: 4, maxItems: 4 },
          correctIndex: { type: Type.INTEGER },
          explanation: { type: Type.STRING }
        },
        required: ["question", "options", "correctIndex", "explanation"]
      }
    },
    qa: {
      type: Type.ARRAY,
      minItems: 10,
      maxItems: 10,
      items: {
        type: Type.OBJECT,
        properties: {
          question: { type: Type.STRING },
          answer: { type: Type.STRING }
        },
        required: ["question", "answer"]
      }
    },
    interactive: {
      type: Type.OBJECT,
      properties: {
        crossword: {
          type: Type.ARRAY,
          minItems: 10,
          maxItems: 10,
          items: {
            type: Type.OBJECT,
            properties: {
              definition: { type: Type.STRING },
              answer: { type: Type.STRING }
            }
          }
        },
        puzzle: { type: Type.STRING },
        game: {
          type: Type.OBJECT,
          properties: {
            name: { type: Type.STRING },
            rules: { type: Type.ARRAY, items: { type: Type.STRING } }
          }
        }
      }
    }
  },
  required: ["slides", "tests", "qa", "interactive"]
};

export async function generateStandaloneImage(prompt: string): Promise<string> {
  const response = await ai.models.generateContent({
    model: 'gemini-2.5-flash-image',
    contents: { parts: [{ text: `High-quality professional 3D educational illustration for a classroom. Topic: ${prompt}. Clean, vibrant, no text, 4K.` }] },
  });
  for (const part of response.candidates[0].content.parts) {
    if (part.inlineData) return `data:image/png;base64,${part.inlineData.data}`;
  }
  throw new Error("Rasm yaratib bo'lmadi");
}

export async function generateLogicPuzzle(topic: string): Promise<{ puzzle: string, solution: string }> {
  const response = await ai.models.generateContent({
    model: 'gemini-3-flash-preview',
    contents: `Mavzu: ${topic}. Maktab o'quvchilari uchun juda qiziqarli va kreativ mantiqiy jumboq yoki topishmoq yarating. Javobini ham alohida yozing.`,
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          puzzle: { type: Type.STRING },
          solution: { type: Type.STRING }
        },
        required: ["puzzle", "solution"]
      }
    }
  });
  return JSON.parse(response.text.trim());
}

export async function generateLesson(params: {
  subject: string;
  grade: string;
  topic: string;
  goal: string;
  language: string;
}): Promise<LessonData> {
  const prompt = `You are an expert educational content creator. Create a complete lesson package in EXACTLY THE ${params.language} LANGUAGE.
    
    CONSTRAINTS:
    - Subject: ${params.subject}
    - Grade: ${params.grade}
    - Topic: ${params.topic}
    - Goal: ${params.goal} (if Revision, focus on key concepts. if Exam, focus on challenging parts. if New Topic, focus on explanation)
    - Language: ${params.language}
    
    OUTPUT:
    1. 10 educational slides with professional content.
    2. 10 interactive tests with 4 options and logical explanations.
    3. 10 Q&A (FAQ) items.
    4. 10 crossword items.
    5. 1 Logic puzzle.
    6. 1 Classroom game.
    
    Ensure all text is high-quality and scientifically accurate.`;

  const response: GenerateContentResponse = await ai.models.generateContent({
    model: 'gemini-3-flash-preview',
    contents: prompt,
    config: {
      responseMimeType: "application/json",
      responseSchema: lessonSchema,
    },
  });

  const parsed = JSON.parse(response.text.trim());

  // Generate images for ALL 10 slides
  const slidesWithImages = await Promise.all(parsed.slides.map(async (slide: any) => {
    try {
      const imgResponse = await ai.models.generateContent({
        model: 'gemini-2.5-flash-image',
        contents: { parts: [{ text: `Professional educational 3D illustration: ${slide.title}. ${slide.imagePrompt}. High definition, vibrant.` }] },
      });
      let imageUrl = '';
      for (const part of imgResponse.candidates[0].content.parts) {
        if (part.inlineData) imageUrl = `data:image/png;base64,${part.inlineData.data}`;
      }
      return { ...slide, imageUrl };
    } catch (e) {
      return slide;
    }
  }));

  return {
    ...parsed,
    slides: slidesWithImages,
    id: Math.random().toString(36).substr(2, 9),
    date: new Date().toLocaleDateString('uz-UZ'),
    ...params
  };
}
