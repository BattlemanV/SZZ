from __future__ import annotations

import argparse
import json

from mr2projekt import settings
from mr2projekt.pipeline import convert, default_output_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Отчёты «Эколога» МР/СГ/СС -> отчёт СЗЗ по шаблону"
    )
    parser.add_argument("--mr", required=True, help="Отчёт МР")
    parser.add_argument("--sg", help="Отчёт СГ (не задан — из МР)")
    parser.add_argument("--ss", help="Отчёт СС (не задан — из МР)")
    parser.add_argument("--template", default=str(settings.template_path()))
    parser.add_argument("--out", help="Результат (по умолчанию — рядом с МР)")
    parser.add_argument(
        "--vocw",
        help="Каталог справочника веществ VocW (по умолчанию — из settings.ini)",
    )
    args = parser.parse_args()

    result = convert(
        args.mr,
        args.template,
        args.out or default_output_path(args.mr),
        annual_path=args.sg,
        daily_path=args.ss,
        vocw_dir=args.vocw,
    )
    print(json.dumps(result.summary, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
