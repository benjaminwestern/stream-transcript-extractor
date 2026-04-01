import { createInterface } from 'node:readline';

const DEFAULT_HELP_WIDTH = 88;

function buildError(message, errorFactory) {
  return errorFactory ? errorFactory(message) : new Error(message);
}

function readOptionValue(flag, inlineValue, nextArg, errorFactory) {
  const value = inlineValue ?? nextArg;

  if (
    value == null ||
    value === '' ||
    (inlineValue == null && String(value).startsWith('--'))
  ) {
    throw buildError(`Missing value for ${flag}.`, errorFactory);
  }

  return value;
}

function wrapText(
  text,
  width = DEFAULT_HELP_WIDTH,
  firstIndent = '',
  restIndent = firstIndent,
) {
  const content = String(text || '').trim();
  if (!content) {
    return [firstIndent.trimEnd()];
  }

  const words = content.split(/\s+/);
  const lines = [];
  let currentIndent = firstIndent;
  let currentContent = '';

  for (const word of words) {
    const candidateContent = currentContent
      ? `${currentContent} ${word}`
      : word;

    if (currentIndent.length + candidateContent.length <= width) {
      currentContent = candidateContent;
      continue;
    }

    if (currentContent) {
      lines.push(`${currentIndent}${currentContent}`);
      currentIndent = restIndent;
      currentContent = word;
      continue;
    }

    lines.push(`${currentIndent}${word}`);
    currentIndent = restIndent;
  }

  if (currentContent) {
    lines.push(`${currentIndent}${currentContent}`);
  }

  return lines;
}

function formatRows(rows) {
  const width = Math.max(
    0,
    ...rows.map((row) => String(row.label || '').length),
  );

  return rows.flatMap((row) => {
    const label = `  ${String(row.label || '').padEnd(width)}`;
    const description = String(row.description || '').trim();
    return wrapText(description, DEFAULT_HELP_WIDTH, `${label}  `, `${' '.repeat(label.length)}  `);
  });
}

export function renderHelpScreen({
  name,
  summary = '',
  usage = [],
  sections = [],
}) {
  const lines = [];

  if (name) {
    lines.push(name);
  }

  if (summary) {
    lines.push(summary);
  }

  if (lines.length > 0) {
    lines.push('');
  }

  if (usage.length > 0) {
    lines.push('Usage:');
    usage.forEach((entry) => {
      lines.push(`  ${entry}`);
    });
  }

  for (const section of sections) {
    if (lines.length > 0 && lines.at(-1) !== '') {
      lines.push('');
    }

    lines.push(`${section.title}:`);

    if (Array.isArray(section.paragraphs)) {
      section.paragraphs.forEach((paragraph) => {
        lines.push(...wrapText(paragraph, DEFAULT_HELP_WIDTH, '  ', '  '));
      });
    }

    if (Array.isArray(section.rows) && section.rows.length > 0) {
      lines.push(...formatRows(section.rows));
    }

    if (Array.isArray(section.lines) && section.lines.length > 0) {
      section.lines.forEach((line) => {
        lines.push(...wrapText(line, DEFAULT_HELP_WIDTH, '  ', '  '));
      });
    }
  }

  return `${lines.join('\n').trimEnd()}\n`;
}

export function printHelpScreen(definition) {
  process.stdout.write(renderHelpScreen(definition));
}

export function createPrompt() {
  const readline = createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return {
    ask(question) {
      return new Promise((resolveAnswer) =>
        readline.question(question, resolveAnswer),
      );
    },
    waitForEnter(message = 'Press Enter to continue...') {
      return new Promise((resolveAnswer) =>
        readline.question(message, () => resolveAnswer()),
      );
    },
    close() {
      readline.close();
    },
  };
}

export function printMenu(title, items) {
  console.log(`\n${title}`);
  items.forEach((item, index) => {
    console.log(`  ${index + 1}. ${item}`);
  });
}

export async function chooseFromList(
  prompt,
  title,
  items,
  renderItem,
  question,
) {
  printMenu(title, items.map(renderItem));

  while (true) {
    const answer = (await prompt.ask(question)).trim();
    const selectedIndex = Number.parseInt(answer, 10) - 1;
    const selectedItem = items[selectedIndex];

    if (selectedItem) {
      return selectedItem;
    }

    console.log('Enter one of the listed numbers.');
  }
}

export function parseSelectionSpec(value, itemCount, errorFactory) {
  const rawValue = String(value || '').trim();

  if (!rawValue) {
    throw buildError('Enter at least one item number or range.', errorFactory);
  }

  if (!Number.isInteger(itemCount) || itemCount < 1) {
    throw buildError('No selectable items are available.', errorFactory);
  }

  const selectedIndexes = new Set();
  const segments = rawValue
    .split(',')
    .map((segment) => segment.trim())
    .filter(Boolean);

  if (segments.length === 0) {
    throw buildError('Enter at least one item number or range.', errorFactory);
  }

  for (const segment of segments) {
    const rangeMatch = segment.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const start = Number.parseInt(rangeMatch[1], 10);
      const end = Number.parseInt(rangeMatch[2], 10);

      if (start < 1 || end < 1 || start > itemCount || end > itemCount) {
        throw buildError(
          `Selection "${segment}" is outside the available range 1-${itemCount}.`,
          errorFactory,
        );
      }

      const lower = Math.min(start, end);
      const upper = Math.max(start, end);
      for (let index = lower; index <= upper; index += 1) {
        selectedIndexes.add(index - 1);
      }
      continue;
    }

    if (!/^\d+$/.test(segment)) {
      throw buildError(
        `Selection "${segment}" is invalid. Use values like 1,3,5-8.`,
        errorFactory,
      );
    }

    const index = Number.parseInt(segment, 10);
    if (index < 1 || index > itemCount) {
      throw buildError(
        `Selection "${segment}" is outside the available range 1-${itemCount}.`,
        errorFactory,
      );
    }

    selectedIndexes.add(index - 1);
  }

  return [...selectedIndexes].sort((left, right) => left - right);
}

export async function chooseManyFromList(
  prompt,
  title,
  items,
  renderItem,
  question,
  errorFactory,
  config = {},
) {
  const { resolveSelection } = config;
  printMenu(title, items.map(renderItem));

  while (true) {
    const answer = (await prompt.ask(question)).trim();

    try {
      const indexes = resolveSelection
        ? resolveSelection(answer, items)
        : parseSelectionSpec(answer, items.length, errorFactory);
      return {
        indexes,
        items: indexes.map((index) => items[index]),
      };
    } catch (error) {
      console.log(
        error instanceof Error ? error.message : 'Enter a valid selection.',
      );
    }
  }
}

function assignOptionValue(options, descriptor, value) {
  if (typeof descriptor.set === 'function') {
    descriptor.set(options, value);
    return;
  }

  if (!descriptor.key) {
    throw new Error(`Option ${descriptor.flag} is missing a target key.`);
  }

  options[descriptor.key] = value;
}

export function parseFlagArgs(
  argv,
  { defaults = {}, options = [], errorFactory } = {},
) {
  const parsed = { ...defaults };
  const optionMap = new Map();

  for (const descriptor of options) {
    optionMap.set(descriptor.flag, descriptor);
    for (const alias of descriptor.aliases || []) {
      optionMap.set(alias, descriptor);
    }
  }

  for (let index = 0; index < argv.length; index += 1) {
    const arg = String(argv[index] || '');
    const [flag, inlineValue] = arg.split(/=(.*)/s, 2);
    const descriptor = optionMap.get(flag);

    if (!descriptor) {
      throw buildError(`Unknown option "${arg}".`, errorFactory);
    }

    if (descriptor.takesValue === false) {
      assignOptionValue(
        parsed,
        descriptor,
        Object.hasOwn(descriptor, 'value') ? descriptor.value : true,
      );
      continue;
    }

    const rawValue = readOptionValue(flag, inlineValue, argv[index + 1], errorFactory);
    const value =
      typeof descriptor.parse === 'function'
        ? descriptor.parse(rawValue, {
            flag,
            options: parsed,
            errorFactory,
          })
        : rawValue;

    assignOptionValue(parsed, descriptor, value);

    if (typeof descriptor.validate === 'function') {
      descriptor.validate(value, parsed, errorFactory);
    }

    if (inlineValue == null) {
      index += 1;
    }
  }

  return parsed;
}

export function parseCliArgs(argv, config) {
  const {
    defaults,
    supportedBrowserKeys,
    supportedOutputFormats,
    normalizeBrowserKey,
    errorFactory,
    extraOptions = [],
  } = config;

  const options = { ...defaults };
  const optionMap = new Map(
    extraOptions.map((option) => [option.flag, option]),
  );

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === '--help' || arg === '-h') {
      options.help = true;
      continue;
    }

    if (arg === '--version' || arg === '-v') {
      options.version = true;
      continue;
    }

    if (arg === '--keep-browser-open') {
      options.keepBrowserOpen = true;
      continue;
    }

    if (arg === '--debug') {
      options.debug = true;
      continue;
    }

    if (!arg.startsWith('--')) {
      throw buildError(`Unexpected positional argument "${arg}".`, errorFactory);
    }

    const [flag, inlineValue] = arg.split(/=(.*)/s, 2);
    const value = readOptionValue(flag, inlineValue, argv[index + 1], errorFactory);

    if (flag === '--browser') {
      options.browser = normalizeBrowserKey(value);
    } else if (flag === '--profile') {
      options.profile = value;
    } else if (flag === '--output') {
      options.outputName = value;
    } else if (flag === '--output-dir') {
      options.outputDir = value;
    } else if (flag === '--format') {
      options.outputFormat = String(value).trim().toLowerCase();
    } else if (flag === '--debug-port') {
      const parsedPort = Number.parseInt(value, 10);
      if (!Number.isInteger(parsedPort) || parsedPort < 1 || parsedPort > 65_535) {
        throw buildError(`Invalid port "${value}".`, errorFactory);
      }
      options.debugPort = parsedPort;
    } else if (optionMap.has(flag)) {
      const option = optionMap.get(flag);
      options[option.key] = option.normalize ? option.normalize(value) : value;
    } else {
      throw buildError(`Unknown option "${flag}".`, errorFactory);
    }

    if (inlineValue == null) {
      index += 1;
    }
  }

  if (options.browser && !supportedBrowserKeys.includes(options.browser)) {
    throw buildError(
      `Unsupported browser "${options.browser}". Use one of: ` +
        `${supportedBrowserKeys.join(', ')}.`,
      errorFactory,
    );
  }

  if (!supportedOutputFormats.includes(options.outputFormat)) {
    throw buildError(
      `Unsupported output format "${options.outputFormat}". Use one of: ` +
        `${supportedOutputFormats.join(', ')}.`,
      errorFactory,
    );
  }

  for (const option of extraOptions) {
    if (typeof option.validate === 'function') {
      option.validate(options[option.key], options, errorFactory);
    }
  }

  return options;
}
