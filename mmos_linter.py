#!/usr/bin/env python3
"""
Markdown Linter - Microsoft Manual of Style Edition
Enforces Microsoft's style guide for technical documentation
"""

import re
import sys
from pathlib import Path
from typing import List, Dict, Set, Tuple
from dataclasses import dataclass
from enum import Enum


class Severity(Enum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


@dataclass
class LintIssue:
    line_number: int
    column: int
    severity: Severity
    rule: str
    message: str
    line_content: str
    suggestion: str = ""


class MicrosoftStyleLinter:
    """Markdown linter following Microsoft Manual of Style"""
    
    def __init__(self):
        self.issues: List[LintIssue] = []
        self.init_style_rules()
    
    def init_style_rules(self):
        """Initialize Microsoft style guide rules"""
        
        # Forbidden words and phrases (use simple/clear alternatives)
        self.forbidden_phrases = {
            "please note": "Note:",
            "please remember": "Remember:",
            "in order to": "to",
            "it is recommended": "we recommend",
            "utilize": "use",
            "terminate": "end",
            "initiate": "start",
            "leverage": "use",
            "click on": "select",
            "simply": "",
            "just": "",
            "easy": "",
            "easily": "",
            "obviously": "",
            "clearly": "",
            "basically": "",
            "actually": "",
            "very": "",
            "via": "through",
            "e.g.": "for example",
            "i.e.": "that is",
            "etc.": "and so on",
            "as well as": "and",
            "shall": "must" or "will",
            "may or may not": "might",
            "in the event that": "if",
            "at this point in time": "now",
            "for the purpose of": "to",
            "due to the fact that": "because",
        }
        
        # Contractions to avoid in technical documentation
        self.contractions = [
            "don't", "won't", "can't", "shouldn't", "wouldn't", 
            "couldn't", "isn't", "aren't", "wasn't", "weren't",
            "haven't", "hasn't", "hadn't", "it's", "that's",
            "there's", "here's", "what's", "who's", "they're"
        ]
        
        # Latin abbreviations to avoid
        self.latin_abbrev = {
            "e.g.": "for example",
            "i.e.": "that is",
            "etc.": "and so on",
            "et al.": "and others",
            "via": "through",
            "vs.": "versus",
            "cf.": "compare",
        }
        
        # Preferred Microsoft terminology
        self.preferred_terms = {
            "webpage": "web page",
            "web site": "website",
            "internet": "Internet",
            "email": "email",  # not e-mail
            "log in": "sign in",  # verb
            "login": "sign-in",  # noun/adjective
            "log on": "sign in",
            "logon": "sign-in",
            "logout": "sign out",
            "username": "user name",
            "wifi": "Wi-Fi",
            "backend": "back end",
            "frontend": "front end",
            "enduser": "end user",
            "file name": "filename",  # Microsoft prefers one word
            "check box": "checkbox",  # Microsoft prefers one word
            "double click": "double-click",
            "right click": "right-click",
            "grayed": "dimmed" or "unavailable",
            "grayed out": "unavailable",
            "hit": "press" or "select",
            "type in": "enter" or "type",
            "box": "dialog box" or "text box",
            "deselect": "clear",
            "uncheck": "clear",
        }
        
        # Avoid directional language
        self.directional_words = [
            "above", "below", "left", "right", "top", "bottom",
            "upper", "lower", "here", "there"
        ]
        
        # Avoid future tense
        self.future_indicators = [
            "will be", "will have", "will see", "is going to",
            "are going to", "will allow", "will enable"
        ]
        
        # Gender-neutral alternatives
        self.gendered_terms = {
            "he": "they",
            "she": "they",
            "his": "their",
            "her": "their",
            "him": "them",
            "himself": "themselves",
            "herself": "themselves",
            "mankind": "people" or "humanity",
            "manmade": "synthetic" or "artificial",
            "man-hours": "person-hours" or "hours",
        }
        
        # Common technical misspellings per Microsoft
        self.spelling_corrections = {
            "cancelled": "canceled",
            "cancelling": "canceling",
            "labelled": "labeled",
            "labelling": "labeling",
            "modelling": "modeling",
            "travelled": "traveled",
            "focussed": "focused",
            "workaround": "work around",  # verb
            "work-around": "workaround",  # noun
        }
        
        # Phrases requiring specific capitalization
        self.capitalization_rules = {
            "internet": "Internet",
            "web": "web",  # lowercase unless starting sentence
            "cloud": "cloud",  # lowercase unless part of product name
        }
        
        # Punctuation rules
        self.punctuation_rules = {
            "oxford_comma": True,
            "serial_comma": True,
            "no_double_spaces": True,
            "space_after_period": True,
        }
    
    def lint_file(self, filepath: str) -> List[LintIssue]:
        """Lint a markdown file and return issues found"""
        self.issues = []
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
                lines = content.splitlines()
        except Exception as e:
            self.add_issue(0, 0, Severity.ERROR, "file-read", 
                          f"Error reading file: {e}", "", "")
            return self.issues
        
        # Run all Microsoft style rules
        self.check_forbidden_phrases(lines)
        self.check_contractions(lines)
        self.check_latin_abbreviations(lines)
        self.check_preferred_terminology(lines)
        self.check_directional_language(lines)
        self.check_future_tense(lines)
        self.check_gender_neutral(lines)
        self.check_voice(lines)
        self.check_headings_microsoft(lines)
        self.check_lists_microsoft(lines)
        self.check_capitalization(lines)
        self.check_punctuation(lines)
        self.check_numbers_and_units(lines)
        self.check_ui_elements(lines)
        self.check_code_formatting(lines)
        self.check_links_microsoft(lines)
        self.check_accessibility(lines)
        self.check_inclusive_language(lines)
        
        # Sort issues by line number
        self.issues.sort(key=lambda x: (x.line_number, x.column))
        return self.issues
    
    def add_issue(self, line_num: int, col: int, severity: Severity,
                  rule: str, message: str, line_content: str, suggestion: str = ""):
        """Add a lint issue with optional suggestion"""
        self.issues.append(LintIssue(
            line_number=line_num,
            column=col,
            severity=severity,
            rule=rule,
            message=message,
            line_content=line_content,
            suggestion=suggestion
        ))
    
    def is_in_code_block(self, lines: List[str], line_index: int) -> bool:
        """Check if a line is inside a code block"""
        in_code = False
        for i in range(line_index):
            if re.match(r'^```|^~~~', lines[i]):
                in_code = not in_code
        return in_code
    
    def check_forbidden_phrases(self, lines: List[str]):
        """Check for forbidden phrases per Microsoft style"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            line_lower = line.lower()
            for phrase, replacement in self.forbidden_phrases.items():
                if phrase in line_lower:
                    col = line_lower.index(phrase)
                    suggestion = f"Replace with: {replacement}" if replacement else "Remove this word"
                    self.add_issue(i, col, Severity.WARNING, "ms-forbidden-phrase",
                                 f"Avoid '{phrase}'", line, suggestion)
    
    def check_contractions(self, lines: List[str]):
        """Check for contractions (avoid in technical docs)"""
        contraction_pattern = re.compile(r"\b(" + "|".join(map(re.escape, self.contractions)) + r")\b", re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            matches = contraction_pattern.finditer(line)
            for match in matches:
                contraction = match.group(1).lower()
                # Provide expanded form
                expanded = {
                    "don't": "do not", "won't": "will not", "can't": "cannot",
                    "shouldn't": "should not", "wouldn't": "would not",
                    "couldn't": "could not", "isn't": "is not", "aren't": "are not",
                    "wasn't": "was not", "weren't": "were not", "haven't": "have not",
                    "hasn't": "has not", "hadn't": "had not", "it's": "it is",
                    "that's": "that is", "there's": "there is", "here's": "here is",
                    "what's": "what is", "who's": "who is", "they're": "they are"
                }
                suggestion = f"Use: {expanded.get(contraction, 'expanded form')}"
                self.add_issue(i, match.start(), Severity.WARNING, "ms-contraction",
                             f"Avoid contractions in technical documentation: '{match.group(1)}'",
                             line, suggestion)
    
    def check_latin_abbreviations(self, lines: List[str]):
        """Check for Latin abbreviations"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            for abbrev, replacement in self.latin_abbrev.items():
                if abbrev in line.lower():
                    col = line.lower().index(abbrev)
                    self.add_issue(i, col, Severity.WARNING, "ms-latin-abbrev",
                                 f"Avoid Latin abbreviation '{abbrev}'", line,
                                 f"Use: {replacement}")
    
    def check_preferred_terminology(self, lines: List[str]):
        """Check for Microsoft preferred terminology"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            line_lower = line.lower()
            for incorrect, correct in self.preferred_terms.items():
                pattern = re.compile(r'\b' + re.escape(incorrect) + r'\b', re.IGNORECASE)
                matches = pattern.finditer(line)
                for match in matches:
                    self.add_issue(i, match.start(), Severity.INFO, "ms-preferred-term",
                                 f"Microsoft prefers '{correct}' over '{match.group()}'",
                                 line, f"Use: {correct}")
    
    def check_directional_language(self, lines: List[str]):
        """Check for directional language (accessibility issue)"""
        directional_pattern = re.compile(
            r'\b(above|below|left|right|top|bottom|upper|lower)\b(?!\s+(menu|bar|corner|section|panel))',
            re.IGNORECASE
        )
        
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            # Skip if it's in a code example or referring to UI elements
            if '`' in line or 'screenshot' in line.lower():
                continue
            
            matches = directional_pattern.finditer(line)
            for match in matches:
                self.add_issue(i, match.start(), Severity.WARNING, "ms-directional",
                             f"Avoid directional language: '{match.group()}'", line,
                             "Use 'preceding', 'following', or specific references instead")
    
    def check_future_tense(self, lines: List[str]):
        """Check for future tense (use present tense)"""
        future_pattern = re.compile(
            r'\b(will be|will have|will see|will allow|will enable|is going to|are going to)\b',
            re.IGNORECASE
        )
        
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            matches = future_pattern.finditer(line)
            for match in matches:
                self.add_issue(i, match.start(), Severity.INFO, "ms-future-tense",
                             f"Prefer present tense over future: '{match.group()}'", line,
                             "Use present tense (e.g., 'allows', 'enables', 'displays')")
    
    def check_gender_neutral(self, lines: List[str]):
        """Check for gender-specific language"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            line_lower = line.lower()
            words = re.findall(r'\b\w+\b', line_lower)
            
            for word in words:
                if word in self.gendered_terms:
                    col = line_lower.index(word)
                    self.add_issue(i, col, Severity.WARNING, "ms-gender-neutral",
                                 f"Use gender-neutral language: '{word}'", line,
                                 f"Consider: {self.gendered_terms[word]}")
    
    def check_voice(self, lines: List[str]):
        """Check for passive voice (prefer active)"""
        # Common passive voice indicators
        passive_pattern = re.compile(
            r'\b(is|are|was|were|be|been|being)\s+\w+ed\b',
            re.IGNORECASE
        )
        
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            matches = passive_pattern.finditer(line)
            for match in matches:
                # Avoid false positives
                if not any(word in line.lower() for word in ['speed', 'need', 'feed']):
                    self.add_issue(i, match.start(), Severity.INFO, "ms-passive-voice",
                                 "Consider using active voice", line,
                                 "Rewrite in active voice for clarity")
    
    def check_headings_microsoft(self, lines: List[str]):
        """Check heading formatting per Microsoft style"""
        heading_pattern = re.compile(r'^(#{1,6})\s+(.+)$')
        
        for i, line in enumerate(lines, 1):
            match = heading_pattern.match(line)
            if match:
                level = len(match.group(1))
                title = match.group(2).strip()
                
                # Check sentence case (Microsoft prefers sentence case)
                words = title.split()
                if len(words) > 1:
                    # Check if it's title case when it should be sentence case
                    title_case_count = sum(1 for w in words[1:] if w[0].isupper() and len(w) > 3)
                    if title_case_count > len(words[1:]) / 2:
                        self.add_issue(i, 0, Severity.INFO, "ms-heading-case",
                                     "Microsoft prefers sentence case for headings", line,
                                     "Use sentence case (capitalize only first word and proper nouns)")
                
                # Check for ending punctuation (should not have)
                if title.rstrip().endswith(('.', '!', '?', ':')):
                    self.add_issue(i, len(line)-1, Severity.WARNING, "ms-heading-punctuation",
                                 "Headings should not end with punctuation", line,
                                 "Remove trailing punctuation")
                
                # Check for gerunds in headings (prefer infinitives or nouns)
                first_word = words[0] if words else ""
                if first_word.lower().endswith('ing'):
                    self.add_issue(i, 0, Severity.INFO, "ms-heading-gerund",
                                 f"Consider using imperative or noun form instead of gerund: '{first_word}'",
                                 line, "Use imperative (e.g., 'Configure') or noun form")
    
    def check_lists_microsoft(self, lines: List[str]):
        """Check list formatting per Microsoft style"""
        list_pattern = re.compile(r'^(\s*)([-*+]|\d+\.)\s+(.+)$')
        
        for i, line in enumerate(lines, 1):
            match = list_pattern.match(line)
            if match:
                content = match.group(3)
                
                # Check parallel structure
                if content and content[0].islower():
                    # Check if previous list item exists and compare
                    if i > 1:
                        prev_match = list_pattern.match(lines[i-2])
                        if prev_match:
                            prev_content = prev_match.group(3)
                            if prev_content and prev_content[0].isupper():
                                self.add_issue(i, 0, Severity.INFO, "ms-list-parallel",
                                             "List items should have parallel structure", line,
                                             "Ensure all list items start with same case")
                
                # Check ending punctuation
                # Complete sentences should end with period
                if len(content.split()) > 5 and not content.rstrip().endswith(('.', ':', '!')):
                    self.add_issue(i, len(line)-1, Severity.INFO, "ms-list-punctuation",
                                 "Complete sentences in lists should end with period", line,
                                 "Add period if this is a complete sentence")
    
    def check_capitalization(self, lines: List[str]):
        """Check capitalization rules"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            # Check "Internet" capitalization
            internet_pattern = re.compile(r'\binternet\b')
            matches = internet_pattern.finditer(line)
            for match in matches:
                self.add_issue(i, match.start(), Severity.WARNING, "ms-internet-cap",
                             "'Internet' should be capitalized", line,
                             "Use: Internet")
    
    def check_punctuation(self, lines: List[str]):
        """Check punctuation rules"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            # Check for double spaces
            if '  ' in line:
                col = line.index('  ')
                self.add_issue(i, col, Severity.INFO, "ms-double-space",
                             "Use single space between sentences", line,
                             "Remove extra space")
            
            # Check for spaces before punctuation
            if re.search(r'\s+[.,!?;:]', line):
                match = re.search(r'\s+[.,!?;:]', line)
                self.add_issue(i, match.start(), Severity.WARNING, "ms-space-before-punct",
                             "No space before punctuation", line,
                             "Remove space before punctuation mark")
            
            # Check for Oxford comma in lists
            list_pattern = re.compile(r'\w+,\s+\w+\s+and\s+\w+')
            if list_pattern.search(line):
                match = list_pattern.search(line)
                self.add_issue(i, match.start(), Severity.INFO, "ms-oxford-comma",
                             "Microsoft style uses serial comma", line,
                             "Add comma before 'and' in series")
    
    def check_numbers_and_units(self, lines: List[str]):
        """Check number and unit formatting"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            # Check for spelled out numbers at start of sentence
            number_start = re.match(r'^(\d+)', line.strip())
            if number_start:
                self.add_issue(i, 0, Severity.INFO, "ms-number-start",
                             "Spell out numbers at the beginning of a sentence", line,
                             "Rewrite sentence or spell out the number")
            
            # Check for space between number and unit
            no_space_pattern = re.compile(r'\d+(GB|MB|KB|TB|ms|cm|mm|km|kg|mg)')
            matches = no_space_pattern.finditer(line)
            for match in matches:
                self.add_issue(i, match.start(), Severity.WARNING, "ms-number-unit-space",
                             "Add space between number and unit", line,
                             f"Use: {match.group().rstrip('GMKTBmcskbg')} {match.group()[-2:]}")
            
            # Check for zero before decimal
            decimal_pattern = re.compile(r'(?<!\d)\.\d+')
            matches = decimal_pattern.finditer(line)
            for match in matches:
                self.add_issue(i, match.start(), Severity.INFO, "ms-zero-decimal",
                             "Include zero before decimal point", line,
                             f"Use: 0{match.group()}")
    
    def check_ui_elements(self, lines: List[str]):
        """Check UI element formatting"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            # Check for unformatted UI elements
            ui_verbs = ['click', 'select', 'choose', 'press', 'enter', 'type']
            for verb in ui_verbs:
                pattern = re.compile(rf'\b{verb}\s+(?!the\s+)`?\w+`?', re.IGNORECASE)
                matches = pattern.finditer(line)
                for match in matches:
                    # Check if the following word is in bold or code format
                    after_verb = line[match.end():match.end()+20]
                    if not (after_verb.startswith('**') or after_verb.startswith('`')):
                        if verb.lower() in ['click', 'press', 'select']:
                            preferred = 'select'
                            self.add_issue(i, match.start(), Severity.INFO, "ms-ui-element",
                                         f"UI elements should be bold. Also prefer '{preferred}' over '{verb}'",
                                         line, "Format UI element in bold (**element**)")
            
            # Check for "click on" - should be just "select"
            if re.search(r'\bclick on\b', line, re.IGNORECASE):
                match = re.search(r'\bclick on\b', line, re.IGNORECASE)
                self.add_issue(i, match.start(), Severity.WARNING, "ms-click-on",
                             "Use 'select' instead of 'click on'", line,
                             "Replace 'click on' with 'select'")
    
    def check_code_formatting(self, lines: List[str]):
        """Check code formatting"""
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            # Check for unformatted code elements
            code_indicators = ['function', 'method', 'class', 'variable', 'parameter', 'property']
            for indicator in code_indicators:
                pattern = re.compile(rf'\b{indicator}\s+\w+\b(?!`)', re.IGNORECASE)
                matches = pattern.finditer(line)
                for match in matches:
                    # Check if the word after indicator is in code format
                    word_after = re.search(r'\w+', line[match.end():])
                    if word_after:
                        pos = match.end() + word_after.start()
                        if pos < len(line) and line[pos-1] != '`':
                            self.add_issue(i, pos, Severity.INFO, "ms-code-format",
                                         "Code elements should be formatted as code", line,
                                         "Wrap in backticks (`code`)")
    
    def check_links_microsoft(self, lines: List[str]):
        """Check link formatting per Microsoft style"""
        inline_link = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')
        
        for i, line in enumerate(lines, 1):
            matches = inline_link.finditer(line)
            for match in matches:
                link_text = match.group(1)
                url = match.group(2)
                
                # Check for "click here" or "here" as link text
                if link_text.lower() in ['click here', 'here', 'this link', 'this page']:
                    self.add_issue(i, match.start(), Severity.WARNING, "ms-link-text",
                                 f"Avoid '{link_text}' as link text", line,
                                 "Use descriptive link text that makes sense out of context")
                
                # Check for URL as link text
                if url.startswith('http') and link_text.startswith('http'):
                    self.add_issue(i, match.start(), Severity.INFO, "ms-link-descriptive",
                                 "Use descriptive link text instead of URL", line,
                                 "Replace URL with meaningful description")
                
                # Check for "see" or "refer to" before link
                before_link = line[:match.start()].lower().strip()
                if before_link.endswith(('see', 'see the', 'refer to', 'refer to the')):
                    self.add_issue(i, match.start()-10, Severity.INFO, "ms-link-see",
                                 "Link text should be descriptive without 'see' or 'refer to'",
                                 line, "Remove 'see' or 'refer to' and use descriptive link text")
    
    def check_accessibility(self, lines: List[str]):
        """Check for accessibility issues"""
        for i, line in enumerate(lines, 1):
            # Check for images without alt text
            img_pattern = re.compile(r'!\[\s*\]\([^)]+\)')
            matches = img_pattern.finditer(line)
            for match in matches:
                self.add_issue(i, match.start(), Severity.ERROR, "ms-alt-text",
                             "Images must have alt text for accessibility", line,
                             "Add descriptive alt text: ![description](url)")
            
            # Check for color-only indicators
            color_only = ['red', 'green', 'blue', 'yellow', 'colored', 'highlighted']
            line_lower = line.lower()
            for color in color_only:
                pattern = re.compile(rf'\b{color}\b(?!\s+(text|background|icon|button))')
                if pattern.search(line_lower):
                    match = pattern.search(line_lower)
                    self.add_issue(i, match.start(), Severity.WARNING, "ms-color-only",
                                 f"Don't rely solely on color ('{color}') to convey information",
                                 line, "Add text labels or icons in addition to color")
    
    def check_inclusive_language(self, lines: List[str]):
        """Check for inclusive language (avoiding ableist, racist terms)"""
        
        # Terms to avoid with alternatives
        problematic_terms = {
            "blacklist": "blocklist or denylist",
            "whitelist": "allowlist or safelist",
            "master": "primary or main",
            "slave": "secondary or replica",
            "native": "built-in or core",
            "grandfathered": "legacy",
            "dummy": "placeholder or sample",
            "sanity check": "consistency check or validation",
            "crazy": "unexpected or surprising",
            "insane": "extreme or intensive",
            "cripple": "disable or limit",
            "blind": "without knowledge of or hidden",
            "man-in-the-middle": "machine-in-the-middle or interceptor",
        }
        
        for i, line in enumerate(lines, 1):
            if self.is_in_code_block(lines, i-1):
                continue
            
            line_lower = line.lower()
            for term, alternative in problematic_terms.items():
                pattern = re.compile(rf'\b{re.escape(term)}\b', re.IGNORECASE)
                matches = pattern.finditer(line)
                for match in matches:
                    self.add_issue(i, match.start(), Severity.WARNING, "ms-inclusive",
                                 f"Use inclusive language: avoid '{match.group()}'", line,
                                 f"Consider: {alternative}")


def format_issue(issue: LintIssue, show_line: bool = True, use_color: bool = True) -> str:
    """Format a lint issue for display"""
    if use_color:
        severity_colors = {
            Severity.ERROR: "\033[91m",  # Red
            Severity.WARNING: "\033[93m",  # Yellow
            Severity.INFO: "\033[94m",  # Blue
        }
        reset = "\033[0m"
        bold = "\033[1m"
    else:
        severity_colors = {s: "" for s in Severity}
        reset = ""
        bold = ""
    
    color = severity_colors.get(issue.severity, "")
    output = f"{color}{bold}{issue.severity.value.upper()}{reset}: "
    output += f"Line {issue.line_number}:{issue.column} "
    output += f"{bold}[{issue.rule}]{reset} {issue.message}"
    
    if show_line and issue.line_content:
        output += f"\n  {color}│{reset} {issue.line_content}"
    
    if issue.suggestion:
        output += f"\n  {color}└─{reset} 💡 {issue.suggestion}"
    
    return output


def generate_report(issues: List[LintIssue], filepath: str) -> Dict:
    """Generate a summary report"""
    errors = sum(1 for i in issues if i.severity == Severity.ERROR)
    warnings = sum(1 for i in issues if i.severity == Severity.WARNING)
    info = sum(1 for i in issues if i.severity == Severity.INFO)
    
    rule_counts = {}
    for issue in issues:
        rule_counts[issue.rule] = rule_counts.get(issue.rule, 0) + 1
    
    return {
        'filepath': filepath,
        'total_issues': len(issues),
        'errors': errors,
        'warnings': warnings,
        'info': info,
        'rule_counts': rule_counts
    }


def main():
    """CLI entry point"""
    import argparse
    import json
    
    parser = argparse.ArgumentParser(
        description="Markdown linter based on Microsoft Manual of Style",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s README.md
  %(prog)s docs/*.md --severity warning
  %(prog)s --json report.json docs/
  %(prog)s --fix README.md (auto-fix where possible)

For more information on Microsoft style guide:
https://docs.microsoft.com/en-us/style-guide/
        """
    )
    
    parser.add_argument('files', nargs='+', help='Markdown files to lint')
    parser.add_argument('--severity', choices=['error', 'warning', 'info'],
                       default='info', help='Minimum severity to report')
    parser.add_argument('--no-color', action='store_true', help='Disable color output')
    parser.add_argument('--show-lines', action='store_true', default=True,
                       help='Show line content for issues')
    parser.add_argument('--json', metavar='FILE', help='Output results as JSON')
    parser.add_argument('--summary', action='store_true', help='Show summary only')
    parser.add_argument('--group-by-rule', action='store_true',
                       help='Group issues by rule')
    
    args = parser.parse_args()
    
    linter = MicrosoftStyleLinter()
    severity_levels = {'error': 3, 'warning': 2, 'info': 1}
    min_severity = severity_levels[args.severity]
    
    all_results = []
    total_issues = 0
    exit_code = 0
    
    for filepath in args.files:
        path = Path(filepath)
        
        if not path.exists():
            print(f"Error: File not found: {filepath}")
            continue
        
        if path.is_dir():
            md_files = list(path.rglob('*.md'))
        else:
            md_files = [path]
        
        for md_file in md_files:
            issues = linter.lint_file(str(md_file))
            
            # Filter by severity
            filtered_issues = [
                issue for issue in issues
                if severity_levels[issue.severity.value] >= min_severity
            ]
            
            report = generate_report(filtered_issues, str(md_file))
            all_results.append(report)
            
            if not args.summary:
                print(f"\n{'='*70}")
                print(f"📄 Linting: {md_file}")
                print(f"{'='*70}")
                
                if not filtered_issues:
                    print("✅ No issues found - Microsoft style compliant!")
                else:
                    if args.group_by_rule:
                        # Group issues by rule
                        by_rule = {}
                        for issue in filtered_issues:
                            if issue.rule not in by_rule:
                                by_rule[issue.rule] = []
                            by_rule[issue.rule].append(issue)
                        
                        for rule, rule_issues in sorted(by_rule.items()):
                            print(f"\n📋 Rule: {rule} ({len(rule_issues)} issues)")
                            print("─" * 70)
                            for issue in rule_issues:
                                print(format_issue(issue, args.show_lines, not args.no_color))
                    else:
                        for issue in filtered_issues:
                            print(format_issue(issue, args.show_lines, not args.no_color))
                    
                    if any(i.severity == Severity.ERROR for i in filtered_issues):
                        exit_code = 1
                
                total_issues += len(filtered_issues)
    
    # Print summary
    print(f"\n{'='*70}")
    print("📊 SUMMARY")
    print(f"{'='*70}")
    
    total_errors = sum(r['errors'] for r in all_results)
    total_warnings = sum(r['warnings'] for r in all_results)
    total_info = sum(r['info'] for r in all_results)
    
    print(f"Files checked: {len(all_results)}")
    print(f"Total issues: {total_issues}")
    print(f"  ❌ Errors: {total_errors}")
    print(f"  ⚠️  Warnings: {total_warnings}")
    print(f"  ℹ️  Info: {total_info}")
    
    # Top issues
    if all_results:
        all_rule_counts = {}
        for result in all_results:
            for rule, count in result['rule_counts'].items():
                all_rule_counts[rule] = all_rule_counts.get(rule, 0) + count
        
        if all_rule_counts:
            print(f"\n🔝 Top Issues:")
            for rule, count in sorted(all_rule_counts.items(), 
                                     key=lambda x: x[1], reverse=True)[:5]:
                print(f"  • {rule}: {count}")
    
    # JSON output
    if args.json:
        with open(args.json, 'w') as f:
            json.dump({
                'summary': {
                    'total_files': len(all_results),
                    'total_issues': total_issues,
                    'errors': total_errors,
                    'warnings': total_warnings,
                    'info': total_info
                },
                'files': all_results
            }, f, indent=2)
        print(f"\n📄 JSON report saved to: {args.json}")
    
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
