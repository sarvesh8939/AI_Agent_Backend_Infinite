import { ChatGoogleGenerativeAI } from "@langchain/google-genai";
import { createReactAgent } from "@langchain/langgraph/prebuilt";
import { DynamicTool } from "@langchain/core/tools";
import axios from "axios";
//import { TavilySearch } from "@langchain/tavily";
//import fs from "fs";
//import path from "path";
//import os from "os";
import ExcelJS from "exceljs";
import nodemailer from "nodemailer";
import dns from "dns";

// Force Node.js to use IPv4 instead of IPv6 to fix Render's ENETUNREACH error
dns.setDefaultResultOrder("ipv4first");

// 1. Define the Internet Search Tool (Uses Tavily for reliable searching)
async function performInternetSearch(query) {
    try {
        const response = await axios.post('https://api.tavily.com/search', {
            api_key: process.env.TAVILY_API_KEY,
            query: query,
            max_results: 5
        });
        
        if (!response.data || !response.data.results || response.data.results.length === 0) {
            return "No search results found for: " + query;
        }

        const results = response.data.results.map(r => 
            `Title: ${r.title}\nSnippet: ${r.content}\nURL: ${r.url}`
        );
        return results.join('\n\n');
    } catch (error) {
        return `Failed to search the internet using Tavily: ${error.message}`;
    }
}

const searchInternetTool = new DynamicTool({
    name: "search_internet",
    description: "Useful for searching the internet to find competitors, market research, trends, and general information. Input should be a precise search query.",
    func: performInternetSearch
});

// 2. Define Generate Excel Tool
const generateExcelTool = new DynamicTool({
    name: "generate_excel",
    description: "Generates an Excel spreadsheet from the compiled business research data. Input MUST be a valid JSON string containing keys like 'businessIdea', 'competitors', 'marketTrends', 'topKeywordsforSEO', 'targetAudience', 'challenges', 'UniqueSellingProposition', and 'confidence'. Returns the absolute file path to the generated Excel file. Do NOT pass markdown inside the JSON string.",
    func: async (jsonString) => {
        try {
            let cleanString = jsonString.trim();
            if (cleanString.startsWith('```json')) cleanString = cleanString.replace(/```json\n?/, '').replace(/```$/, '').trim();
            else if (cleanString.startsWith('```')) cleanString = cleanString.replace(/```\n?/, '').replace(/```$/, '').trim();
            
            const parsedData = JSON.parse(cleanString);
            const workbook = new ExcelJS.Workbook();
            
            // Sheet 1: General Info
            const overviewSheet = workbook.addWorksheet("Overview");
            overviewSheet.columns = [
                { header: 'Business Idea', key: 'idea', width: 50 },
                { header: 'Unique Selling Proposition', key: 'usp', width: 50 },
                { header: 'Confidence', key: 'confidence', width: 15 },
            ];
            overviewSheet.addRow({
                idea: parsedData.businessIdea,
                usp: parsedData.UniqueSellingProposition,
                confidence: parsedData.confidence
            });

            // Sheet 2: Competitors
            const compSheet = workbook.addWorksheet("Competitors");
            compSheet.columns = [
                { header: 'Name', key: 'name', width: 25 },
                { header: 'Description', key: 'description', width: 60 },
                { header: 'URL', key: 'url', width: 35 },
            ];
            if (parsedData.competitors) {
                parsedData.competitors.forEach(comp => compSheet.addRow(comp));
            }

            // Sheet 3: Market Trends & Keywords
            const trendsSheet = workbook.addWorksheet("Trends & Keywords");
            trendsSheet.columns = [
                { header: 'Market Trends', key: 'trend', width: 40 },
                { header: 'SEO Keyword', key: 'keyword', width: 30 },
                { header: 'Volume', key: 'volume', width: 15 },
            ];
            
            const maxRows = Math.max(
                parsedData.marketTrends?.length || 0, 
                parsedData.topKeywordsforSEO?.length || 0
            );
            for (let i = 0; i < maxRows; i++) {
                trendsSheet.addRow({
                    trend: parsedData.marketTrends?.[i] || "",
                    keyword: parsedData.topKeywordsforSEO?.[i]?.name || "",
                    volume: parsedData.topKeywordsforSEO?.[i]?.searchVolume || ""
                });
            }

            const fileName = `Research_Report_${Date.now()}.xlsx`;
            const filePath = `./${fileName}`; // Relative path
            await workbook.xlsx.writeFile(filePath);
            
            return filePath;
        } catch (error) {
            return `Failed to generate Excel: ${error.message}`;
        }
    }
});

// 3. Define Send Email Tool
const sendEmailTool = new DynamicTool({
    name: "send_email",
    description: "Sends the generated Excel report to the user's email. Input MUST be a valid JSON string with 'email' (the user's email address) and 'attachmentPath' (the file path returned by generate_excel). Example: {\"email\": \"user@example.com\", \"attachmentPath\": \"C:/path/to/Research_Report_123.xlsx\"}",
    func: async (inputStr) => {
        try {
            let cleanString = inputStr.trim();
            if (cleanString.startsWith('```json')) cleanString = cleanString.replace(/```json\n?/, '').replace(/```$/, '').trim();
            else if (cleanString.startsWith('```')) cleanString = cleanString.replace(/```\n?/, '').replace(/```$/, '').trim();
            
            const { email, attachmentPath } = JSON.parse(cleanString);
            
            if (!email || !attachmentPath) {
                return "Error: Both 'email' and 'attachmentPath' must be provided in the JSON.";
            }

            if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
                return "Error: EMAIL_USER or EMAIL_PASS environment variables are not set on the server.";
            }

            let transporter = nodemailer.createTransport({
                host: 'smtp.gmail.com',
                port: 465,
                secure: true, // Use port 465 with secure: true for Render
                auth: {
                    user: process.env.EMAIL_USER,
                    pass: process.env.EMAIL_PASS
                }
            });

            let mailOptions = {
                from: `"AI Research Agent" <${process.env.EMAIL_USER}>`,
                to: email,
                subject: 'Your AI Business Research Report',
                text: 'Hello! Your AI business research is complete. Please find the attached Excel report.',
                attachments: [
                    { path: attachmentPath }
                ]
            };

            await transporter.sendMail(mailOptions);
            return `Email sent successfully to ${email}.`;
        } catch (error) {
            console.error("--- NODEMAILER ERROR ---", error);
            return `Failed to send email: ${error.message}`;
        }
    }
});

// 4. Define Tools array
const tools = [searchInternetTool, generateExcelTool, sendEmailTool];

const agentInstructions = `You are a professional business research AI agent. 
When the user provides a business idea, title, or requirements, you MUST autonomously research the internet to find:
1. Top Competitors in that space.
2. Current Market Trends.
3. Other relevant details (target audience, challenges, etc.).

Use the 'search_internet' tool to query for competitors, market trends, and specific competitor details.

Once you have gathered all the research data:
1. Format ALL the gathered data into a single JSON object matching the final answer structure below.
2. Call the 'generate_excel' tool, passing the ENTIRE JSON object as a string. This tool will return a file path.
3. If an email address is provided by the user, call the 'send_email' tool, passing a JSON string containing the user's 'email' and the 'attachmentPath' returned by the generate_excel tool.

Always output your final answer as a raw JSON object with the following structure (representing your research findings):
{
  "businessIdea": "Summarized business idea",
  "competitors": [
    { "name": "Competitor 1", "description": "What they do", "url": "their URL if found" }
  ],
  "marketTrends": ["Trend 1", "Trend 2"],
  "topKeywordsforSEO": [{"name":"keyword1","searchVolume":"10k"},{"name":"keyword2","searchVolume":"1.5m"}], // Ensure searchVolume is formatted with k, m, or b (e.g., 10k, 1.5m, 2b)
  "targetAudience": {
    "primary": "Primary target audience",
    "secondary": "Secondary target audience" 
  },
  "challenges": ["Challenge 1", "Challenge 2"],
  "UniqueSellingProposition": "What makes your business unique and valuable",
  "confidence": "High/Medium/Low",
  "actionsTaken": ["Generated Excel", "Sent Email to user@example.com"]
}

Do NOT wrap your JSON in markdown code blocks (like \`\`\`json). Just output the raw JSON text.`;

export async function runAgent(req, res) {
    // The user input text and optionally an email address
    const { message, email } = req.body;

    if (!message) {
        return res.status(400).json({ error: "Please provide a 'message' in the request body." });
    }

    

    if (!process.env.GEMINI_API_KEY) {
        return res.status(400).json({ error: "GEMINI_API_KEY is not set in the .env file." });
    }

    if (!process.env.TAVILY_API_KEY) {
        return res.status(400).json({ error: "TAVILY_API_KEY is not set in the .env file. Required for internet research." });
    }

    try {
        // Initialize Model inside the function to prevent startup crashes if key is missing
        const model = new ChatGoogleGenerativeAI({
            apiKey: process.env.GEMINI_API_KEY,
            model: "gemini-2.5-flash",
            temperature: 0.2,  
            maxRetries: 5, // Automatically wait and retry on 429 rate limit errors
        }); 

        // Create the Agent dynamically
        const agent = createReactAgent({
            llm: model,
            tools,
            stateModifier: agentInstructions
        });

        // Inject email into the user message if provided
        let finalMessage = message;
        if (email) {
            finalMessage += `\n\nUser's email address is: ${email}. Make sure to send the final report to this email by using the send_email tool.`;
        }

        // Invoke the Agent (This will trigger the internal reasoning loop)
        const response = await agent.invoke({
            messages: [{ role: "user", content: finalMessage }],
        });

        // The last message in the array is the agent's final conclusion
        const finalAnswer = response.messages[response.messages.length - 1].content;
        
        let parsedData;
        try {
            // Attempt to parse the JSON
            let cleanString = finalAnswer.trim();
            if (cleanString.startsWith('```json')) {
                cleanString = cleanString.replace(/```json\n?/, '').replace(/```$/, '').trim();
            } else if (cleanString.startsWith('```')) {
                cleanString = cleanString.replace(/```\n?/, '').replace(/```$/, '').trim();
            }
            parsedData = JSON.parse(cleanString);
        } catch (e) {
            // Fallback to text if the agent fails to output valid JSON
            parsedData = finalAnswer; 
        }

        res.json({ success: true, data: parsedData });

    } catch (error) {
        console.error("Error in agent:", error);
        res.status(500).json({ success: false, error: error.message });
    }
};
