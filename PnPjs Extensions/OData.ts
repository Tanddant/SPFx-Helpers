type KeysMatching<T, V> = { [K in keyof T]-?: T[K] extends V ? K : never }[keyof T];

enum FilterOperation {
    Equals = "eq",
    NotEquals = "ne",
    GreaterThan = "gt",
    GreaterThanOrEqualTo = "ge",
    LessThan = "lt",
    LessThanOrEqualTo = "le",
    StartsWith = "startswith",
    EndsWith = "endswith",
    SubstringOf = "substringof"
}

enum FilterJoinOperator {
    And = "and",
    AndWithSpace = " and ",
    Or = "or",
    OrWithSpace = " or "
}

export namespace OData {
    export const Where = <T>() => new ODataFilterClass<T>();

    class BaseQueryable<T> {
        protected query: string[] = [];

        constructor(query: string[]) {
            this.query = query;
        }
    }

    class WhereClause<T> extends BaseQueryable<T> {
        constructor(q: string[]) {
            super(q);
        }

        public TextField(InternalName: KeysMatching<T, string>): TextField<T> {
            return new TextField<T>([...this.query, (InternalName as string)]);
        }

        public NumberField(InternalName: KeysMatching<T, number>): NumberField<T> {
            return new NumberField<T>([...this.query, (InternalName as string)]);
        }

        public DateField(InternalName: KeysMatching<T, Date>): DateField<T> {
            return new DateField<T>([...this.query, (InternalName as string)]);
        }

        public BooleanField(InternalName: KeysMatching<T, boolean>): BooleanField<T> {
            return new BooleanField<T>([...this.query, (InternalName as string)]);
        }
    }

    class ODataFilterClass<T> extends WhereClause<T>{
        constructor() {
            super([]);
        }

        public All(queries: BaseFilterCompareResult<T>[]): BaseFilterCompareResult<T> {
            return new BaseFilterCompareResult<T>(["(", queries.map(x => x.ToString()).join(FilterJoinOperator.AndWithSpace), ")"]);
        }

        public Some(queries: BaseFilterCompareResult<T>[]): BaseFilterCompareResult<T> {
            return new BaseFilterCompareResult<T>(["(", queries.map(x => x.ToString()).join(FilterJoinOperator.OrWithSpace), ")"]);
        }
    }

    class BaseField<Ttype, Tinput> extends BaseQueryable<Ttype>{
        constructor(q: string[]) {
            super(q);
        }

        protected ToODataValue(value: Tinput): string {
            return `'${value}'`;
        }

        public EqualTo(value: Tinput): BaseFilterCompareResult<Ttype> {
            return new BaseFilterCompareResult<Ttype>([...this.query, FilterOperation.Equals, this.ToODataValue(value)]);
        }

        public NotEqualTo(value: Tinput): BaseFilterCompareResult<Ttype> {
            return new BaseFilterCompareResult<Ttype>([...this.query, FilterOperation.NotEquals, this.ToODataValue(value)]);
        }

        public IsNull(): BaseFilterCompareResult<Ttype> {
            return new BaseFilterCompareResult<Ttype>([...this.query, FilterOperation.Equals, "null"]);
        }

        public IsNotNull(): BaseFilterCompareResult<Ttype> {
            return new BaseFilterCompareResult<Ttype>([...this.query, FilterOperation.NotEquals, "null"]);
        }
    }

    class BaseComparableField<Tinput, Ttype> extends BaseField<Tinput, Ttype>{
        constructor(q: string[]) {
            super(q);
        }

        public GreaterThan(value: Ttype): BaseFilterCompareResult<Tinput> {
            return new BaseFilterCompareResult<Tinput>([...this.query, FilterOperation.GreaterThan, this.ToODataValue(value)]);
        }

        public GreaterThanOrEqualTo(value: Ttype): BaseFilterCompareResult<Tinput> {
            return new BaseFilterCompareResult<Tinput>([...this.query, FilterOperation.GreaterThanOrEqualTo, this.ToODataValue(value)]);
        }

        public LessThan(value: Ttype): BaseFilterCompareResult<Tinput> {
            return new BaseFilterCompareResult<Tinput>([...this.query, FilterOperation.LessThan, this.ToODataValue(value)]);
        }

        public LessThanOrEqualTo(value: Ttype): BaseFilterCompareResult<Tinput> {
            return new BaseFilterCompareResult<Tinput>([...this.query, FilterOperation.LessThanOrEqualTo, this.ToODataValue(value)]);
        }
    }

    class TextField<T> extends BaseField<T, string>{
        constructor(q: string[]) {
            super(q);
        }
    }

    class NumberField<T> extends BaseComparableField<T, number>{
        constructor(q: string[]) {
            super(q);
        }

        protected override ToODataValue(value: number): string {
            return `${value}`;
        }
    }

    class DateField<T> extends BaseComparableField<T, Date>{
        constructor(q: string[]) {
            super(q);
        }

        protected override ToODataValue(value: Date): string {
            return `'${value.toISOString()}'`
        }
    }

    class BooleanField<T> extends BaseField<T, boolean>{
        constructor(q: string[]) {
            super(q);
        }

        protected override ToODataValue(value: boolean): string {
            return `${value == null ? "null" : value ? 1 : 0}`;
        }
    }


    class BaseFilterCompareResult<T> extends BaseQueryable<T>{
        constructor(q: string[]) {
            super(q);
        }

        public Or(): FilterOr<T> {
            return new FilterOr<T>(this.query);
        }

        public And(): FilterAnd<T> {
            return new FilterAnd<T>(this.query);
        }

        public ToString(): string {
            return this.query.join(" ");
        }
    }

    class FilterOr<T> extends WhereClause<T>{
        constructor(currentQuery: string[]) {
            super([...currentQuery, FilterJoinOperator.Or]);
        }
    }

    class FilterAnd<T> extends WhereClause<T>{
        constructor(currentQuery: string[]) {
            super([...currentQuery, FilterJoinOperator.And]);
        }
    }
}