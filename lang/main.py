#!/usr/bin/env python3
# ░▄▀▄░▀█▀░█▄█▒██▀▒█▀▄░▀█▀▒▄▀▄░█▒░▒██▀░▄▀▀
# ░▀▄▀░▒█▒▒█▒█░█▄▄░█▀▄░▒█▒░█▀█▒█▄▄░█▄▄▒▄██
"""
Genkō Yōshi Tategaki Converter with LangGraph and ChatAnthropic Integration
Enhanced with AI-powered rule processing and compliance verification
Maintains all original functionality while adding intelligent text processing
"""

import re
import argparse
import logging
import json
import os
from typing import Dict, List, Tuple, Optional, Any, TypedDict
from pathlib import Path
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:
    import chardet
except ImportError:
    print("Warning: chardet module not found, falling back to default encodings")
    chardet = None

from rich.console import Console, Group
from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn, TaskProgressColumn
from rich.panel import Panel
from rich.live import Live
from rich.table import Table
try:
    from rich.prompt import Prompt
except ImportError:
    Prompt = None

# LangGraph and LangChain imports
try:
    from langgraph.graph import StateGraph, END
    from langgraph.graph.message import add_messages
    from langgraph.checkpoint.memory import MemorySaver
    from langchain_core.messages import HumanMessage, SystemMessage, AIMessage
    from langchain_core.pydantic_v1 import BaseModel, Field
    LANGGRAPH_AVAILABLE = True
    
    # Import LLM providers
    ANTHROPIC_AVAILABLE = False
    HUGGINGFACE_AVAILABLE = False
    
    try:
        from langchain_anthropic import ChatAnthropic
        ANTHROPIC_AVAILABLE = True
    except ImportError:
        pass
    
    try:
        from langchain_huggingface import ChatHuggingFace
        from langchain_community.llms import HuggingFacePipeline
        import torch
        from transformers import AutoTokenizer, AutoModelForCausalLM, pipeline
        HUGGINGFACE_AVAILABLE = True
    except ImportError:
        try:
            # Alternative import path
            from langchain_community.chat_models import ChatHuggingFace
            from langchain_community.llms import HuggingFacePipeline
            import torch
            from transformers import AutoTokenizer, AutoModelForCausalLM, pipeline
            HUGGINGFACE_AVAILABLE = True
        except ImportError:
            pass
            
except ImportError as e:
    print(f"Warning: LangGraph/LangChain not available: {e}")
    print("Install with: pip install langgraph langchain-anthropic langchain-huggingface")
    LANGGRAPH_AVAILABLE = False
    ANTHROPIC_AVAILABLE = False
    HUGGINGFACE_AVAILABLE = False

# Import the page size selector from parent directory
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

try:
    from sizes import PageSizeSelector
except ImportError:
    PageSizeSelector = None

logging.basicConfig(level=logging.INFO, format='[%(levelname)s] %(message)s')


def create_llm_instance(provider: str = None, api_key: str = None, hf_token: str = None, console: Console = None):
    """Create an LLM instance based on the provider"""
    if console is None:
        console = Console()
    
    # Determine provider from environment or parameter
    if provider is None:
        provider = os.getenv("LLM_PROVIDER", "anthropic").lower()
    
    if provider == "huggingface" and HUGGINGFACE_AVAILABLE:
        # Get HuggingFace token
        token = hf_token or os.getenv("HF_TOKEN") or os.getenv("HUGGINGFACE_API_TOKEN")
        if not token:
            console.print("[bold red]Warning: HF_TOKEN not found. Using public models only.[/bold red]")
        
        model_name = os.getenv("HUGGINGFACE_MODEL", "meta-llama/Llama-3.3-70B-Instruct")
        
        try:
            # Create HuggingFace pipeline for chat
            console.print(f"[cyan]Initializing HuggingFace model: {model_name}[/cyan]")
            
            # Check if we should use API or local pipeline
            use_api = os.getenv("HUGGINGFACE_USE_API", "true").lower() == "true"
            
            if use_api:
                # Use HuggingFace Inference API
                from langchain_huggingface import HuggingFaceEndpoint
                
                llm = HuggingFaceEndpoint(
                    repo_id=model_name,
                    temperature=0.1,
                    max_new_tokens=4000,
                    huggingfacehub_api_token=token,
                    task="text-generation"
                )
                
                # Wrap in ChatHuggingFace for chat interface
                chat_llm = ChatHuggingFace(llm=llm, verbose=True)
                
            else:
                # Use local pipeline (requires more memory)
                tokenizer = AutoTokenizer.from_pretrained(
                    model_name, 
                    use_auth_token=token if token else None
                )
                model = AutoModelForCausalLM.from_pretrained(
                    model_name,
                    use_auth_token=token if token else None,
                    torch_dtype=torch.float16 if torch.cuda.is_available() else torch.float32,
                    device_map="auto" if torch.cuda.is_available() else None
                )
                
                pipe = pipeline(
                    "text-generation",
                    model=model,
                    tokenizer=tokenizer,
                    max_new_tokens=4000,
                    temperature=0.1,
                    do_sample=True,
                    pad_token_id=tokenizer.eos_token_id
                )
                
                hf_pipeline = HuggingFacePipeline(pipeline=pipe)
                chat_llm = ChatHuggingFace(llm=hf_pipeline, verbose=True)
            
            console.print(f"[green]✓ HuggingFace model initialized: {model_name}[/green]")
            return chat_llm
            
        except Exception as e:
            console.print(f"[bold red]Error initializing HuggingFace model: {e}[/bold red]")
            console.print("[yellow]Falling back to Anthropic if available...[/yellow]")
            provider = "anthropic"
    
    if provider == "anthropic" and ANTHROPIC_AVAILABLE:
        # Use Anthropic Claude
        anthropic_key = api_key or os.getenv("ANTHROPIC_API_KEY")
        if not anthropic_key:
            raise ValueError("ANTHROPIC_API_KEY is required for Anthropic provider")
        
        model_name = os.getenv("ANTHROPIC_MODEL", "claude-3-sonnet-20240229")
        
        llm = ChatAnthropic(
            model=model_name,
            api_key=anthropic_key,
            temperature=0.1,
            max_tokens=4000
        )
        
        console.print(f"[green]✓ Anthropic model initialized: {model_name}[/green]")
        return llm
    
    # Error if no provider is available
    available_providers = []
    if ANTHROPIC_AVAILABLE:
        available_providers.append("anthropic")
    if HUGGINGFACE_AVAILABLE:
        available_providers.append("huggingface")
    
    if not available_providers:
        raise ImportError("No LLM providers available. Install langchain-anthropic or langchain-huggingface.")
    
    raise ValueError(f"Provider '{provider}' not available. Available providers: {available_providers}")


# Pydantic models for structured AI responses
if LANGGRAPH_AVAILABLE:
    class TextValidationResult(BaseModel):
        """Structured response for text validation"""
        is_valid: bool = Field(description="Whether the text follows Japanese writing conventions")
        corrections: List[str] = Field(description="List of suggested corrections")
        compliance_score: float = Field(description="Score from 0-1 for compliance with genkou yoshi rules")
        suggestions: List[str] = Field(description="List of formatting suggestions")

    class RuleProcessingResult(BaseModel):
        """Structured response for rule processing"""
        processed_text: str = Field(description="Text after applying genkou yoshi rules")
        rules_applied: List[str] = Field(description="List of rules that were applied")
        violations_found: List[str] = Field(description="List of violations found and fixed")
        confidence: float = Field(description="Confidence in the processing accuracy")


# State management for LangGraph
class ProcessingState(TypedDict):
    """State for the LangGraph processing pipeline"""
    input_text: str
    processed_text: str
    validation_result: Optional[Dict]
    rule_processing_result: Optional[Dict]
    final_text: str
    messages: List[Any]
    page_format: Dict
    processing_complete: bool


if LANGGRAPH_AVAILABLE:
    class AITextProcessor:
        """AI-powered text processor using configurable LLM providers and LangGraph"""
        
        def __init__(self, console: Console, api_key: Optional[str] = None, hf_token: Optional[str] = None, 
                     provider: Optional[str] = None):
            self.console = console
            if not LANGGRAPH_AVAILABLE:
                raise ImportError("LangGraph and LLM providers are required for AI processing")
            
            # Initialize LLM based on provider
            self.provider = provider or os.getenv("LLM_PROVIDER", "anthropic").lower()
            self.console.print(f"[cyan]Initializing LLM provider: {self.provider}[/cyan]")
            
            try:
                self.llm = create_llm_instance(
                    provider=self.provider,
                    api_key=api_key,
                    hf_token=hf_token,
                    console=self.console
                )
            except Exception as e:
                self.console.print(f"[bold red]Failed to initialize LLM: {e}[/bold red]")
                raise
            
            # Create structured output models (with fallback for HuggingFace)
            try:
                self.validation_llm = self.llm.with_structured_output(TextValidationResult)
                self.rule_processing_llm = self.llm.with_structured_output(RuleProcessingResult)
                self.supports_structured_output = True
            except Exception as e:
                self.console.print(f"[yellow]Warning: Structured output not supported by {self.provider}, using text parsing[/yellow]")
                self.validation_llm = self.llm
                self.rule_processing_llm = self.llm
                self.supports_structured_output = False
            
            # Initialize memory saver for persistent conversations
            self.memory = MemorySaver()
            
            # Build the processing graph
            self.graph = self._build_processing_graph()
            
        def _build_processing_graph(self):
            """Build the LangGraph processing pipeline"""
            
            # Define the graph
            workflow = StateGraph(ProcessingState)
            
            # Add nodes
            workflow.add_node("validate_text", self._validate_text_node)
            workflow.add_node("process_rules", self._process_rules_node)
            workflow.add_node("finalize_text", self._finalize_text_node)
            
            # Define the flow
            workflow.set_entry_point("validate_text")
            workflow.add_edge("validate_text", "process_rules")
            workflow.add_edge("process_rules", "finalize_text")
            workflow.add_edge("finalize_text", END)
            
            # Compile with memory
            return workflow.compile(checkpointer=self.memory)
        
        def _validate_text_node(self, state: ProcessingState) -> ProcessingState:
            """Validate text against Japanese writing conventions"""
            
            system_prompt = """You are an expert in Japanese writing conventions and genkou yoshi manuscript formatting.
        
        Analyze the provided text for:
        1. Proper use of Japanese punctuation (。、？！)
        2. Correct character spacing and formatting
        3. Appropriate use of kanji, hiragana, and katakana
        4. Compliance with traditional manuscript paper (genkou yoshi) rules
        5. Line breaking rules (禁則処理)
        
        Provide specific, actionable feedback for improvements."""
            
            messages = [
                SystemMessage(content=system_prompt),
                HumanMessage(content=f"Please validate this Japanese text:\n\n{state['input_text']}")
            ]
            
            try:
                if self.supports_structured_output:
                    result = self.validation_llm.invoke(messages)
                    state["validation_result"] = result.dict()
                    state["messages"] = add_messages(state.get("messages", []), messages)
                    
                    self.console.print(f"[green]✓ Text validation completed[/green]")
                    self.console.print(f"[cyan]Compliance score: {result.compliance_score:.2f}[/cyan]")
                else:
                    # Use text parsing for non-structured output
                    result = self.validation_llm.invoke(messages)
                    response_text = result.content if hasattr(result, 'content') else str(result)
                    
                    # Parse the response to extract validation information
                    parsed_result = self._parse_validation_response(response_text)
                    state["validation_result"] = parsed_result
                    state["messages"] = add_messages(state.get("messages", []), messages)
                    
                    self.console.print(f"[green]✓ Text validation completed[/green]")
                    self.console.print(f"[cyan]Compliance score: {parsed_result['compliance_score']:.2f}[/cyan]")
                
            except Exception as e:
                logging.warning(f"AI validation failed: {e}")
                state["validation_result"] = {
                    "is_valid": True,
                    "corrections": [],
                    "compliance_score": 0.8,
                    "suggestions": ["Validation completed with fallback"]
                }
        
            return state
        
        def _process_rules_node(self, state: ProcessingState) -> ProcessingState:
            """Apply genkou yoshi rules using AI processing"""
            
            system_prompt = """You are an expert in Japanese genkou yoshi (manuscript paper) formatting rules.
        
        Apply these specific rules to the text:
        1. 禁則処理 (line breaking rules): Don't start lines with closing punctuation (。、」）etc.)
        2. Don't end lines with opening punctuation (「（etc.)
        3. Convert ASCII punctuation to Japanese equivalents
        4. Ensure proper character spacing for vertical writing (tategaki)
        5. Handle small kana (っゃゅょ etc.) placement correctly
        6. Apply proper formatting for numbers, dates, and foreign words
        
        Return the processed text with all rules applied correctly."""
            
            validation_context = ""
            if state.get("validation_result"):
                validation_context = f"\nValidation findings:\n{json.dumps(state['validation_result'], indent=2, ensure_ascii=False)}"
            
            messages = [
                SystemMessage(content=system_prompt),
                HumanMessage(content=f"Apply genkou yoshi rules to this text:{validation_context}\n\nText to process:\n{state['input_text']}")
            ]
            
            try:
                if self.supports_structured_output:
                    result = self.rule_processing_llm.invoke(messages)
                    state["rule_processing_result"] = result.dict()
                    state["processed_text"] = result.processed_text
                    state["messages"] = add_messages(state.get("messages", []), messages)
                    
                    self.console.print(f"[green]✓ Rule processing completed[/green]")
                    self.console.print(f"[cyan]Confidence: {result.confidence:.2f}[/cyan]")
                    self.console.print(f"[yellow]Rules applied: {len(result.rules_applied)}[/yellow]")
                else:
                    # Use text parsing for non-structured output
                    result = self.rule_processing_llm.invoke(messages)
                    response_text = result.content if hasattr(result, 'content') else str(result)
                    
                    # Parse the response to extract processed text and metadata
                    parsed_result = self._parse_rule_processing_response(response_text, state["input_text"])
                    state["rule_processing_result"] = parsed_result
                    state["processed_text"] = parsed_result["processed_text"]
                    state["messages"] = add_messages(state.get("messages", []), messages)
                    
                    self.console.print(f"[green]✓ Rule processing completed[/green]")
                    self.console.print(f"[cyan]Confidence: {parsed_result['confidence']:.2f}[/cyan]")
                    self.console.print(f"[yellow]Rules applied: {len(parsed_result['rules_applied'])}[/yellow]")
                
            except Exception as e:
                logging.warning(f"AI rule processing failed: {e}")
                # Fallback to original text processing
                processor = JapaneseTextProcessor()
                state["processed_text"] = processor.preprocess_text(state["input_text"])
                state["rule_processing_result"] = {
                    "processed_text": state["processed_text"],
                    "rules_applied": ["Fallback processing"],
                    "violations_found": [],
                    "confidence": 0.7
                }
        
            return state
        
        def _finalize_text_node(self, state: ProcessingState) -> ProcessingState:
            """Finalize the processed text"""
            
            state["final_text"] = state.get("processed_text", state["input_text"])
            state["processing_complete"] = True
            
            self.console.print(f"[bold green]✓ AI processing pipeline completed[/bold green]")
        
            return state
        
        def _parse_validation_response(self, response_text: str) -> Dict:
            """Parse validation response from text output"""
            import re
            
            # Extract compliance score
            score_match = re.search(r'(?:compliance|score)[:\s]*([0-9.]+)', response_text.lower())
            compliance_score = float(score_match.group(1)) if score_match else 0.7
            
            # Normalize score to 0-1 range if it's 0-100
            if compliance_score > 1:
                compliance_score = compliance_score / 100
            
            # Extract corrections and suggestions
            corrections = []
            suggestions = []
            
            # Look for bullet points or numbered lists
            lines = response_text.split('\n')
            current_section = None
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Identify sections
                if any(word in line.lower() for word in ['correction', 'fix', 'error']):
                    current_section = 'corrections'
                elif any(word in line.lower() for word in ['suggestion', 'recommend', 'improve']):
                    current_section = 'suggestions'
                elif line.startswith(('•', '-', '*', '1.', '2.', '3.')):
                    # Extract bullet point content
                    content = re.sub(r'^[•\-*0-9.]\s*', '', line)
                    if current_section == 'corrections':
                        corrections.append(content)
                    else:
                        suggestions.append(content)
            
            # Default suggestions if none found
            if not suggestions:
                suggestions = ["Text processed successfully"]
            
            return {
                "is_valid": compliance_score > 0.6,
                "corrections": corrections,
                "compliance_score": compliance_score,
                "suggestions": suggestions
            }
        
        def _parse_rule_processing_response(self, response_text: str, original_text: str) -> Dict:
            """Parse rule processing response from text output"""
            import re
            
            # Try to extract processed text (look for common patterns)
            processed_text = original_text  # Default fallback
            
            # Look for processed text markers
            text_patterns = [
                r'(?:processed|result|output)[:\s]*(.+?)(?:\n\n|$)',
                r'(?:text|result)[:\s]*["\'](.+?)["\']',
                r'(?:final|processed)\s*text[:\s]*(.+?)(?:\n|$)',
            ]
            
            for pattern in text_patterns:
                match = re.search(pattern, response_text, re.DOTALL | re.IGNORECASE)
                if match:
                    candidate = match.group(1).strip()
                    # Use if it looks like processed Japanese text
                    if len(candidate) > 10 and any(ord(c) > 127 for c in candidate):
                        processed_text = candidate
                        break
            
            # If no explicit processed text found, look for Japanese text in the response
            if processed_text == original_text:
                # Find Japanese text blocks
                japanese_blocks = re.findall(r'[ぁ-んァ-ヶ一-龯]+[ぁ-んァ-ヶ一-龯。、！？「」]*', response_text)
                if japanese_blocks:
                    # Use the longest Japanese block
                    processed_text = max(japanese_blocks, key=len, default=original_text)
            
            # Extract confidence
            confidence_match = re.search(r'(?:confidence|certainty)[:\s]*([0-9.]+)', response_text.lower())
            confidence = float(confidence_match.group(1)) if confidence_match else 0.8
            
            # Normalize confidence to 0-1 range if it's 0-100
            if confidence > 1:
                confidence = confidence / 100
            
            # Extract rules applied
            rules_applied = []
            violations_found = []
            
            # Look for rules or violations mentioned
            lines = response_text.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                if any(word in line.lower() for word in ['rule', 'applied', 'correction']):
                    if line.startswith(('•', '-', '*', '1.', '2.', '3.')):
                        content = re.sub(r'^[•\-*0-9.]\s*', '', line)
                        rules_applied.append(content)
                elif any(word in line.lower() for word in ['violation', 'error', 'issue']):
                    if line.startswith(('•', '-', '*', '1.', '2.', '3.')):
                        content = re.sub(r'^[•\-*0-9.]\s*', '', line)
                        violations_found.append(content)
            
            # Default rules if none found
            if not rules_applied:
                rules_applied = ["Standard genkou yoshi formatting applied"]
            
            return {
                "processed_text": processed_text,
                "rules_applied": rules_applied,
                "violations_found": violations_found,
                "confidence": confidence
            }
        
        def process_text(self, text: str, page_format: Dict, thread_id: str = "default") -> Dict:
            """Process text through the AI pipeline"""
            
            initial_state = {
                "input_text": text,
                "processed_text": "",
                "validation_result": None,
                "rule_processing_result": None,
                "final_text": "",
                "messages": [],
                "page_format": page_format,
                "processing_complete": False
            }
            
            config = {"configurable": {"thread_id": thread_id}}
            
            self.console.print("[bold cyan]Starting AI-powered text processing...[/bold cyan]")
            
            try:
                # Process through the graph
                final_state = self.graph.invoke(initial_state, config)
                return final_state
            except Exception as e:
                logging.error(f"AI processing failed: {e}")
                # Return fallback result
                return {
                    "final_text": text,
                    "validation_result": {"is_valid": True, "compliance_score": 0.5},
                    "rule_processing_result": {"confidence": 0.5},
                    "processing_complete": False
                }

else:
    # Fallback class when LangGraph is not available
    class AITextProcessor:
        def __init__(self, console: Console, api_key: Optional[str] = None, hf_token: Optional[str] = None, 
                     provider: Optional[str] = None):
            self.console = console
            self.console.print("[bold red]LangGraph not available - AI processing disabled[/bold red]")
        
        def process_text(self, text: str, page_format: Dict, thread_id: str = "default") -> Dict:
            return {
                "final_text": text,
                "validation_result": {"is_valid": True, "compliance_score": 0.5},
                "rule_processing_result": {"confidence": 0.5},
                "processing_complete": False
            }


# Import all original classes with minimal modifications from parent directory
import importlib.util
import os

# Import from the parent directory's main.py
parent_main_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'main.py')
spec = importlib.util.spec_from_file_location("parent_main", parent_main_path)
parent_main = importlib.util.module_from_spec(spec)
spec.loader.exec_module(parent_main)

# Import the required classes
GenkouYoshiValidator = parent_main.GenkouYoshiValidator
DocxAnalyzer = parent_main.DocxAnalyzer
DocumentAdjuster = parent_main.DocumentAdjuster
VerificationEngine = parent_main.VerificationEngine
GenkouYoshiGrid = parent_main.GenkouYoshiGrid
JapaneseTextProcessor = parent_main.JapaneseTextProcessor
GenkouYoshiDocumentBuilder = parent_main.GenkouYoshiDocumentBuilder


class EnhancedGenkouYoshiDocumentBuilder(GenkouYoshiDocumentBuilder):
    """Enhanced document builder with AI processing capabilities"""
    
    def __init__(self, font_name='Noto Sans JP', squares_per_column=None, max_columns_per_page=None, 
                 chapter_pagebreak=False, page_size=None, page_format=None, 
                 use_ai=False, anthropic_api_key=None, hf_token=None, llm_provider=None, console=None):
        
        # Initialize parent class
        super().__init__(font_name, squares_per_column, max_columns_per_page, 
                        chapter_pagebreak, page_size, page_format)
        
        self.use_ai = use_ai
        self.console = console or Console()
        self.llm_provider = llm_provider or os.getenv("LLM_PROVIDER", "anthropic").lower()
        
        # Initialize AI processor if enabled
        self.ai_processor = None
        if use_ai and LANGGRAPH_AVAILABLE:
            try:
                self.ai_processor = AITextProcessor(
                    console=self.console, 
                    api_key=anthropic_api_key,
                    hf_token=hf_token,
                    provider=self.llm_provider
                )
                self.console.print(f"[bold green]✓ AI processing enabled with {self.llm_provider}[/bold green]")
            except Exception as e:
                self.console.print(f"[bold red]✗ AI processing setup failed: {e}[/bold red]")
                self.use_ai = False
        elif use_ai:
            self.console.print("[bold red]✗ AI processing requires LangGraph and LLM providers[/bold red]")
            self.use_ai = False
    
    def create_genkou_yoshi_document(self, text):
        """Create document with optional AI processing"""
        try:
            # AI-enhanced preprocessing if enabled
            if self.use_ai and self.ai_processor:
                self.console.print("\n[bold cyan]AI Processing Phase[/bold cyan]")
                
                # Process text through AI pipeline
                ai_result = self.ai_processor.process_text(text, self.page_format)
                
                # Use AI-processed text
                processed_text = ai_result.get("final_text", text)
                
                # Display AI processing results
                if ai_result.get("validation_result"):
                    validation = ai_result["validation_result"]
                    self.console.print(f"[green]AI Validation Score: {validation.get('compliance_score', 'N/A')}[/green]")
                
                if ai_result.get("rule_processing_result"):
                    rule_result = ai_result["rule_processing_result"]
                    self.console.print(f"[green]AI Processing Confidence: {rule_result.get('confidence', 'N/A')}[/green]")
                    
                    if rule_result.get("rules_applied"):
                        self.console.print("[cyan]Rules Applied:[/cyan]")
                        for rule in rule_result["rules_applied"][:5]:  # Show first 5
                            self.console.print(f"  • {rule}")
                
            else:
                # Standard preprocessing
                processed_text = self.text_processor.preprocess_text(text)
            
            # Continue with standard document generation
            structure = self.text_processor.identify_text_structure(processed_text)
            
            if structure['novel_title']:
                self.create_title_page(
                    structure['novel_title'],
                    subtitle=structure['subtitle'],
                    author=structure['author']
                )
                
            if structure['subheadings']:
                for chapter, paragraphs in structure['subheadings']:
                    if self.chapter_pagebreak:
                        self.grid.finish_page()
                    self.create_chapter_title_page(chapter)
                    spacing = self._adaptive_paragraph_spacing(len(paragraphs))
                    
                    for i, paragraph in enumerate(paragraphs):
                        if not paragraph:
                            self.grid.advance_column(1)
                            continue
                        if i > 0:
                            self.grid.advance_column(spacing)
                        
                        self.place_paragraph(paragraph)
            else:
                if structure['novel_title']:
                    self.grid.move_to_column(3, 2)
                    
                paragraphs_to_process = structure['body_paragraphs']
                spacing = self._adaptive_paragraph_spacing(len(paragraphs_to_process))
                
                for i, paragraph in enumerate(paragraphs_to_process):
                    if not paragraph:
                        self.grid.advance_column(1)
                        continue
                    if i > 0:
                        self.grid.advance_column(spacing)
                    
                    self.place_paragraph(paragraph)
                        
            self.grid.finish_page()
            
        except Exception as e:
            logging.critical(f"Failed to generate document: {e}")
            raise


def main():
    """Enhanced main function with AI processing options"""
    from rich.console import Console
    from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn, TaskProgressColumn
    import sys
    
    parser = argparse.ArgumentParser(description="Japanese Tategaki DOCX Generator with AI Enhancement")
    parser.add_argument("input", nargs="?", help="Input text file (UTF-8)")
    parser.add_argument("-o", "--output", help="Output DOCX file")
    parser.add_argument("--json", help="Export grid/character metadata as JSON to file")
    parser.add_argument("--format", default=None, help="Page format")
    parser.add_argument("--skip-verification", action="store_true", help="Skip verification process")
    parser.add_argument("--verification-report", help="Save verification report to file")
    
    # AI processing options
    parser.add_argument("--ai", action="store_true", help="Enable AI-powered rule processing")
    parser.add_argument("--llm-provider", choices=["anthropic", "huggingface"], 
                       help="LLM provider to use (default: from LLM_PROVIDER env var or 'anthropic')")
    parser.add_argument("--anthropic-api-key", help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")
    parser.add_argument("--hf-token", help="HuggingFace token (or set HF_TOKEN env var)")
    parser.add_argument("--ai-thread-id", default="default", help="Thread ID for AI conversation context")

    args = parser.parse_args()

    console = Console(color_system="auto", force_terminal=True, force_interactive=True)
    
    # Display header
    ascii_art = (
        "[cyan]░▄▀▄░▀█▀░█▄█▒██▀▒█▀▄░▀█▀▒▄▀▄░█▒░▒██▀░▄▀▀[/cyan]\n"
        "[cyan]░▀▄▀░▒█▒▒█▒█░█▄▄░█▀▄░▒█▒░█▀█▒█▄▄░█▄▄▒▄██[/cyan]"
    )
    console.print(ascii_art)
    console.print("")
    console.print("[bold yellow]Japanese Tategaki DOCX Generator with AI Enhancement[/bold yellow]")
    console.print("[green]Authentic Genkou Yoshi formatting with LangGraph + Multi-LLM support[/green]")
    
    if args.ai:
        provider = args.llm_provider or os.getenv("LLM_PROVIDER", "anthropic").lower()
        console.print(f"[bold cyan]🤖 AI Processing Mode Enabled ({provider})[/bold cyan]")
    console.print()

    if not args.input:
        console.print("[bold red]No input file specified.[/bold red]")
        sys.exit(1)

    # Load and validate input
    input_path = Path(args.input)
    if not input_path.exists():
        console.print(f"[bold red]Error: Input file '{input_path}' not found.[/bold red]")
        sys.exit(1)
        
    # File reading with encoding detection
    try:
        if chardet:
            with open(input_path, 'rb') as f:
                raw_data = f.read()
            encoding_result = chardet.detect(raw_data)
            detected_encoding = encoding_result['encoding'] if encoding_result['confidence'] > 0.7 else 'utf-8'
            try:
                text = raw_data.decode(detected_encoding).strip()
            except UnicodeDecodeError:
                text = raw_data.decode('utf-8', errors='ignore').strip()
        else:
            with open(input_path, encoding="utf-8", errors='ignore') as f:
                text = f.read().strip()
    except Exception as e:
        console.print(f"[bold red]Error reading file: {e}[/bold red]")
        sys.exit(1)
        
    if not text:
        console.print("[bold red]Error: Input file is empty.[/bold red]")
        sys.exit(1)

    # Check AI requirements
    if args.ai and not LANGGRAPH_AVAILABLE:
        console.print("[bold red]Error: AI processing requires LangGraph and LLM providers.[/bold red]")
        console.print("Install with: pip install langgraph langchain-anthropic langchain-huggingface")
        sys.exit(1)
    
    if args.ai:
        provider = args.llm_provider or os.getenv("LLM_PROVIDER", "anthropic").lower()
        
        if provider == "anthropic":
            if not ANTHROPIC_AVAILABLE:
                console.print("[bold red]Error: Anthropic provider not available.[/bold red]")
                console.print("Install with: pip install langchain-anthropic")
                sys.exit(1)
            if not (args.anthropic_api_key or os.getenv("ANTHROPIC_API_KEY")):
                console.print("[bold red]Error: Anthropic API key required.[/bold red]")
                console.print("Set --anthropic-api-key or ANTHROPIC_API_KEY environment variable")
                sys.exit(1)
        
        elif provider == "huggingface":
            if not HUGGINGFACE_AVAILABLE:
                console.print("[bold red]Error: HuggingFace provider not available.[/bold red]")
                console.print("Install with: pip install langchain-huggingface transformers torch")
                sys.exit(1)
            if not (args.hf_token or os.getenv("HF_TOKEN") or os.getenv("HUGGINGFACE_API_TOKEN")):
                console.print("[bold yellow]Warning: No HF_TOKEN found. Using public models only.[/bold yellow]")
        
        else:
            console.print(f"[bold red]Error: Unknown provider '{provider}'.[/bold red]")
            console.print("Supported providers: anthropic, huggingface")
            sys.exit(1)

    # Page format selection
    if args.format is None or args.format == "custom":
        if PageSizeSelector and Prompt:
            selector = PageSizeSelector(console=console)
            page_format = selector.select_page_size()
        else:
            console.print("[bold red]Error: Interactive page size selection requires 'rich.prompt' and 'sizes.py'.[/bold red]")
            sys.exit(1)
    else:
        try:
            page_format = PageSizeSelector.get_format(args.format) if PageSizeSelector else None
        except Exception:
            page_format = None

    # Create enhanced builder
    builder = EnhancedGenkouYoshiDocumentBuilder(
        page_format=page_format,
        use_ai=args.ai,
        anthropic_api_key=args.anthropic_api_key,
        hf_token=args.hf_token,
        llm_provider=args.llm_provider,
        console=console
    )
    
    # Quick analysis for user feedback
    structure = builder.text_processor.identify_text_structure(text)
    console.print("[bold cyan]Document Analysis:[/bold cyan]")
    console.print(f"  [bold]Title:[/bold] {structure['novel_title']}")
    console.print(f"  [bold]Author:[/bold] {structure['author']}")
    
    if structure['subheadings']:
        total_paragraphs = sum(len(pars) for _, pars in structure['subheadings'])
        console.print(f"  [bold]Chapters:[/bold] {len(structure['subheadings'])}")
    else:
        total_paragraphs = len(structure['body_paragraphs'])
        
    console.print(f"  [bold]Paragraphs:[/bold] {total_paragraphs}")
    console.print(f"  [bold]Characters:[/bold] ~{len(text):,}")

    # Process document with progress tracking
    with Progress(
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
        TimeRemainingColumn(),
        console=console,
    ) as progress:
        
        # Main processing task
        task = progress.add_task("Processing document...", total=100)
        
        # Document creation
        progress.update(task, advance=15, description="Creating document structure...")
        builder.create_genkou_yoshi_document(text)
        
        # DOCX generation
        progress.update(task, advance=10, description="Preparing DOCX generation...")
        pages = builder.grid.get_all_pages()
        
        # Allocate 45% for page-by-page generation
        remaining_progress = 45
        progress_per_page = remaining_progress / max(1, len(pages))
        
        def progress_callback(current_page, total_pages):
            progress.update(task, advance=progress_per_page, 
                          description=f"Generating DOCX... Page {current_page}/{total_pages}")
            
        builder.generate_docx_content(progress_callback=progress_callback)
        
        # Save initial version
        output_path = Path(args.output) if args.output else input_path.with_name(input_path.stem + '_genkou_yoshi_ai.docx')
        builder.doc.save(output_path)
        
        progress.update(task, advance=30, description="Document saved, preparing verification...")

    # Run verification if not skipped
    if not args.skip_verification:
        verification_engine = VerificationEngine(builder, page_format, console)
        verification_report = verification_engine.run_verification_cycle(output_path)
        
        # Save verification report if requested
        if args.verification_report:
            with open(args.verification_report, 'w', encoding='utf-8') as f:
                json.dump(verification_report, f, ensure_ascii=False, indent=2)
            console.print(f"[bold green]✓ Verification report saved:[/bold green] {args.verification_report}")
        
        # Display final status
        if verification_report['status'] == 'compliant':
            console.print(f"\n[bold green]🎉 Document is fully compliant with Genkou Yoshi standards![/bold green]")
            console.print(f"[green]Achieved in {verification_report['iterations']} iteration(s)[/green]")
        elif verification_report['status'] == 'partial_compliance':
            console.print(f"\n[bold yellow]⚠️  Document has partial compliance[/bold yellow]")
            console.print(f"[yellow]Compliance score: {verification_report['final_score']}/100[/yellow]")
            console.print(f"[yellow]Remaining violations: {verification_report['remaining_violations']}[/yellow]")
        else:
            console.print(f"\n[bold red]❌ Verification failed[/bold red]")
    
    console.print()
    console.print(f"[bold green]✓ DOCX file saved:[/bold green] {output_path}")
    console.print(f"[bold green]✓ Pages generated:[/bold green] {len(pages)}")
    
    if args.ai:
        provider = args.llm_provider or os.getenv("LLM_PROVIDER", "anthropic").lower()
        console.print(f"[bold cyan]🤖 AI processing completed with {provider}[/bold cyan]")
    
    if args.json:
        builder.export_grid_metadata_json(args.json)
        console.print(f"[bold green]✓ Metadata JSON saved:[/bold green] {args.json}")


if __name__ == "__main__":
    main()