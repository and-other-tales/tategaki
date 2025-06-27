# Enhanced Tategaki Generator with Multi-LLM Support

This enhanced version of the Japanese Tategaki DOCX Generator adds AI-powered rule processing using LangGraph and supports multiple LLM providers.

## Features

- **LangGraph Integration**: Intelligent text processing pipeline with state management
- **Multi-LLM Support**: Choose between Anthropic Claude or HuggingFace models
- **AI-Powered Rules**: Automated application of genkou yoshi formatting rules
- **Fallback Processing**: Graceful degradation when AI is unavailable
- **All Original Features**: Maintains complete compatibility with the base application

## Installation

Install the enhanced dependencies:

```bash
pip install -r requirements.txt
```

## LLM Provider Configuration

### Environment Variables

Set your preferred LLM provider:

```bash
# Use Anthropic (default)
export LLM_PROVIDER=anthropic
export ANTHROPIC_API_KEY=your_anthropic_key

# Use HuggingFace
export LLM_PROVIDER=huggingface
export HF_TOKEN=your_huggingface_token
```

### HuggingFace Configuration

Additional HuggingFace environment variables:

```bash
# Model selection (default: meta-llama/Llama-3.3-70B-Instruct)
export HUGGINGFACE_MODEL=meta-llama/Llama-3.3-70B-Instruct

# Use API vs local pipeline (default: true)
export HUGGINGFACE_USE_API=true

# Model-specific settings
export ANTHROPIC_MODEL=claude-3-sonnet-20240229
```

## Usage

### Basic Usage (No AI)

```bash
python main.py input.txt --format bunko -o output.docx
```

### With AI Processing

```bash
# Using Anthropic (default)
python main.py input.txt --ai --anthropic-api-key YOUR_KEY -o output.docx

# Using HuggingFace
python main.py input.py --ai --llm-provider huggingface --hf-token YOUR_TOKEN -o output.docx

# Using environment variables
export ANTHROPIC_API_KEY=your_key
python main.py input.txt --ai -o output.docx
```

### Command Line Options

```
--ai                     Enable AI-powered rule processing
--llm-provider {anthropic,huggingface}
                        LLM provider to use
--anthropic-api-key     Anthropic API key
--hf-token              HuggingFace token  
--ai-thread-id          Thread ID for conversation context
```

## AI Processing Pipeline

The LangGraph pipeline consists of three nodes:

1. **Text Validation**: Analyzes Japanese writing conventions and genkou yoshi compliance
2. **Rule Processing**: Applies formatting rules using AI-powered analysis
3. **Text Finalization**: Prepares the processed text for document generation

### Structured vs Unstructured Output

- **Anthropic**: Supports structured output with Pydantic models
- **HuggingFace**: Uses text parsing for response extraction
- **Fallback**: Automatic fallback to original processing when AI fails

## Supported Models

### Anthropic
- claude-3-sonnet-20240229 (default)
- claude-3-opus-20240229
- claude-3-haiku-20240307

### HuggingFace
- meta-llama/Llama-3.3-70B-Instruct (default)
- Any text-generation model with sufficient context length
- Supports both API and local pipeline modes

## API Requirements

### Anthropic
- Valid Anthropic API key
- Sufficient credits for text processing

### HuggingFace
- HuggingFace token (for gated models like Llama)
- API access or local GPU resources
- Model access permissions

## Examples

### Processing with Validation

```bash
python main.py sample.txt --ai --llm-provider anthropic --verification-report report.json
```

### Batch Processing

```bash
for file in *.txt; do
    python main.py "$file" --ai --format bunko -o "${file%.txt}_genkou.docx"
done
```

### Configuration File

Create a `.env` file:

```
LLM_PROVIDER=huggingface
HF_TOKEN=your_token_here
HUGGINGFACE_MODEL=meta-llama/Llama-3.3-70B-Instruct
HUGGINGFACE_USE_API=true
```

## Performance Notes

- **Anthropic**: Fast API responses, structured output
- **HuggingFace API**: Variable response times, depends on model availability
- **HuggingFace Local**: Requires significant GPU memory (70B models need ~140GB VRAM)
- **Fallback Mode**: Instant processing using original algorithms

## Troubleshooting

### Import Errors
```bash
# Install missing dependencies
pip install langgraph langchain-anthropic langchain-huggingface transformers torch
```

### API Authentication
- Verify API keys are correctly set
- Check model access permissions for gated models
- Ensure sufficient credits/quota

### Memory Issues (HuggingFace Local)
- Use smaller models or API mode
- Enable model sharding with `device_map="auto"`
- Consider using quantized models

## License

Same as the original project. Enhanced features maintain compatibility with existing codebase.