
export enum View {
  Loading = 'LOADING',
  Auth = 'AUTH',
  Landing = 'LANDING',
  Dashboard = 'DASHBOARD',
  CreateForm = 'CREATE_FORM',
  Generating = 'GENERATING',
  Results = 'RESULTS',
  AiImage = 'AI_IMAGE',
  LogicGame = 'LOGIC_GAME',
  ViewHistoryItem = 'VIEW_HISTORY_ITEM'
}

export enum AuthMode {
  Login = 'LOGIN',
  Register = 'REGISTER'
}

export enum ResultTab {
  Presentation = 'TAQDIMOT',
  Test = 'INTERAKTIV TEST',
  QA = 'SAVOL-JAVOB',
  Interactive = 'INTERAKTIV'
}

export type HistoryType = 'LESSON' | 'IMAGE' | 'PUZZLE';

export interface HistoryItem {
  id: string;
  type: HistoryType;
  date: string;
  title: string;
  data: any; // Can be LessonData, string (image url), or puzzle object
}

export interface LessonData {
  id: string;
  subject: string;
  grade: string;
  topic: string;
  language: string;
  goal: string;
  date: string;
  slides: Array<{
    title: string;
    content: string;
    imageUrl?: string;
  }>;
  tests: Array<{
    question: string;
    options: string[];
    correctIndex: number;
    explanation: string;
  }>;
  qa: Array<{
    question: string;
    answer: string;
  }>;
  interactive: {
    crossword: Array<{ definition: string; answer: string }>;
    puzzle: string;
    game: { name: string; rules: string[] };
  };
}

export interface User {
  name: string;
  email: string;
  role: string;
  isPremium: boolean;
}
