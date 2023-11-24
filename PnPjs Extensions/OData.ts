type KeysMatching<T, V> = { [K in keyof T]-?: T[K] extends V ? K : never }[keyof T];

enum FilterOperation {
    Equals = "eq",
    NotEquals = "ne",
    GreaterThan = "gt",
    GreaterThanOrEqualTo = "ge",
    LessThan = "lt",
    LessThanOrEqualTo = "le",
    StartsWith = "startswith",
    SubstringOf = "substringof"
}

enum FilterJoinOperator {
    And = "and",
    AndWithSpace = " and ",
    Or = "or",
    OrWithSpace = " or "
}

export class SPOData {
    static Where<T = any>() {
        return new QueryableGroups<T>();
    }
}

class BaseQuery<TBaseInterface> {
    protected query: string[] = [];

    protected AddToQuery(InternalName: keyof TBaseInterface, Operation: FilterOperation, Value: string) {
        this.query.push(`${InternalName as string} ${Operation} ${Value}`);
    }

    protected AddQueryableToQuery(Queries: ComparisonResult<TBaseInterface>) {
        this.query.push(Queries.ToString());
    }

    constructor(BaseQuery?: BaseQuery<TBaseInterface>) {
        if (BaseQuery != null) {
            this.query = BaseQuery.query;
        }
    }
}


class QueryableFields<TBaseInterface> extends BaseQuery<TBaseInterface> {
    public TextField(InternalName: KeysMatching<TBaseInterface, string>): TextField<TBaseInterface> {
        return new TextField<TBaseInterface>(this, InternalName);
    }

    public ChoiceField(InternalName: KeysMatching<TBaseInterface, string>): TextField<TBaseInterface> {
        return new TextField<TBaseInterface>(this, InternalName);
    }

    public MultiChoiceField(InternalName: KeysMatching<TBaseInterface, string[]>): TextField<TBaseInterface> {
        return new TextField<TBaseInterface>(this, InternalName);
    }

    public NumberField(InternalName: KeysMatching<TBaseInterface, number>): NumberField<TBaseInterface> {
        return new NumberField<TBaseInterface>(this, InternalName);
    }

    public DateField(InternalName: KeysMatching<TBaseInterface, Date>): DateField<TBaseInterface> {
        return new DateField<TBaseInterface>(this, InternalName);
    }

    public BooleanField(InternalName: KeysMatching<TBaseInterface, boolean>): BooleanField<TBaseInterface> {
        return new BooleanField<TBaseInterface>(this, InternalName);
    }

    public LookupField<TKey extends KeysMatching<TBaseInterface, object>>(InternalName: TKey): LookupQueryableFields<TBaseInterface, TBaseInterface[TKey]> {
        return new LookupQueryableFields<TBaseInterface, TBaseInterface[TKey]>(this, InternalName as string);
    }

    public LookupIdField<TKey extends KeysMatching<TBaseInterface, number>>(InternalName: TKey): NumberField<TBaseInterface> {
        const col: string = (InternalName as string).endsWith("Id") ? InternalName as string : `${InternalName as string}Id`;
        return new NumberField<TBaseInterface>(this, col as any as keyof TBaseInterface);
    }
}

class LookupQueryableFields<TBaseInterface, TExpandedType> extends BaseQuery<TBaseInterface>{
    private LookupField: string;
    constructor(q: BaseQuery<TBaseInterface>, LookupField: string) {
        super(q);
        this.LookupField = LookupField;
    }

    public Id(Id: number): ComparisonResult<TBaseInterface> {
        this.AddToQuery(`${this.LookupField}Id` as keyof TBaseInterface, FilterOperation.Equals, Id.toString());
        return new ComparisonResult<TBaseInterface>(this);
    }

    public TextField(InternalName: KeysMatching<TExpandedType, string>): TextField<TBaseInterface> {
        return new TextField<TBaseInterface>(this, `${this.LookupField}/${InternalName as string}` as any as keyof TBaseInterface);
    }

    public NumberField(InternalName: KeysMatching<TExpandedType, number>): NumberField<TBaseInterface> {
        return new NumberField<TBaseInterface>(this, `${this.LookupField}/${InternalName as string}` as any as keyof TBaseInterface);
    }

    // Support has been announced, but is not yet available in SharePoint Online
    // https://www.microsoft.com/en-ww/microsoft-365/roadmap?filters=&searchterms=100503
    // public BooleanField(InternalName: KeysMatching<tExpandedType, boolean>): BooleanField<tBaseObjectType> {
    //     return new BooleanField<tBaseObjectType>([...this.query, `${this.LookupField}/${InternalName as string}`]);
    // }
}

class QueryableGroups<TBaseInterface> extends QueryableFields<TBaseInterface>{
    public All(queries: ComparisonResult<TBaseInterface>[]): ComparisonResult<TBaseInterface> {
        this.query.push(`(${queries.map(x => x.ToString()).join(FilterJoinOperator.AndWithSpace)})`);
        return new ComparisonResult<TBaseInterface>(this);
    }

    public Some(queries: ComparisonResult<TBaseInterface>[]): ComparisonResult<TBaseInterface> {
        this.query.push(`(${queries.map(x => x.ToString()).join(FilterJoinOperator.OrWithSpace)})`);
        return new ComparisonResult<TBaseInterface>(this);
    }
}





class NullableField<TBaseInterface, TInputValueType> extends BaseQuery<TBaseInterface>{
    protected InternalName: KeysMatching<TBaseInterface, TInputValueType>;

    constructor(base: BaseQuery<TBaseInterface>, InternalName: keyof TBaseInterface) {
        super(base);
        this.InternalName = InternalName as any as KeysMatching<TBaseInterface, TInputValueType>;
    }

    protected ToODataValue(value: TInputValueType): string {
        return `'${value}'`;
    }

    public IsNull(): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.Equals, "null");
        return new ComparisonResult<TBaseInterface>(this);
    }

    public IsNotNull(): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.NotEquals, "null");
        return new ComparisonResult<TBaseInterface>(this);
    }
}

class ComparableField<TBaseInterface, TInputValueType> extends NullableField<TBaseInterface, TInputValueType>{
    public EqualTo(value: TInputValueType): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.Equals, this.ToODataValue(value));
        return new ComparisonResult<TBaseInterface>(this);
    }

    public NotEqualTo(value: TInputValueType): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.NotEquals, this.ToODataValue(value));
        return new ComparisonResult<TBaseInterface>(this);
    }

    public In(values: TInputValueType[]): ComparisonResult<TBaseInterface> {

        const query = values.map(x =>
            `${this.InternalName as string} ${FilterOperation.Equals} ${this.ToODataValue(x)}`
        ).join(FilterJoinOperator.OrWithSpace);

        this.query.push(`(${query})`);
        return new ComparisonResult<TBaseInterface>(this);
    }
}

class TextField<TBaseInterface> extends ComparableField<TBaseInterface, string>{

    public StartsWith(value: string): ComparisonResult<TBaseInterface> {
        this.query.push(`${FilterOperation.StartsWith}(${this.InternalName as string}, ${this.ToODataValue(value)})`);
        return new ComparisonResult<TBaseInterface>(this);
    }

    public Contains(value: string): ComparisonResult<TBaseInterface> {
        this.query.push(`${FilterOperation.SubstringOf}(${this.ToODataValue(value)}, ${this.InternalName as string})`);
        return new ComparisonResult<TBaseInterface>(this);
    }
}

class BooleanField<TBaseInterface> extends NullableField<TBaseInterface, boolean>{

    protected override ToODataValue(value: boolean | null): string {
        return `${value == null ? "null" : value ? 1 : 0}`;
    }

    public IsTrue(): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.Equals, this.ToODataValue(true));
        return new ComparisonResult<TBaseInterface>(this);
    }

    public IsFalse(): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.Equals, this.ToODataValue(false));
        return new ComparisonResult<TBaseInterface>(this);
    }

    public IsFalseOrNull(): ComparisonResult<TBaseInterface> {
        this.AddQueryableToQuery(SPOData.Where<TBaseInterface>().Some([
            SPOData.Where<TBaseInterface>().BooleanField(this.InternalName).IsFalse(),
            SPOData.Where<TBaseInterface>().BooleanField(this.InternalName).IsNull()
        ]));

        return new ComparisonResult<TBaseInterface>(this);
    }
}

class NumericField<TBaseInterface, TInputValueType> extends ComparableField<TBaseInterface, TInputValueType>{

    public GreaterThan(value: TInputValueType): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.GreaterThan, this.ToODataValue(value));
        return new ComparisonResult<TBaseInterface>(this);
    }

    public GreaterThanOrEqualTo(value: TInputValueType): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.GreaterThanOrEqualTo, this.ToODataValue(value));
        return new ComparisonResult<TBaseInterface>(this);
    }

    public LessThan(value: TInputValueType): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.LessThan, this.ToODataValue(value));
        return new ComparisonResult<TBaseInterface>(this);
    }

    public LessThanOrEqualTo(value: TInputValueType): ComparisonResult<TBaseInterface> {
        this.AddToQuery(this.InternalName, FilterOperation.LessThanOrEqualTo, this.ToODataValue(value));
        return new ComparisonResult<TBaseInterface>(this);
    }
}


class NumberField<TBaseInterface> extends NumericField<TBaseInterface, number>{
    protected override ToODataValue(value: number): string {
        return `${value}`;
    }
}

class DateField<TBaseInterface> extends NumericField<TBaseInterface, Date>{
    protected override ToODataValue(value: Date): string {
        return `'${value.toISOString()}'`
    }

    public IsBetween(startDate: Date, endDate: Date): ComparisonResult<TBaseInterface> {
        this.AddQueryableToQuery(SPOData.Where().All([
            SPOData.Where().DateField(this.InternalName as string).GreaterThanOrEqualTo(startDate),
            SPOData.Where().DateField(this.InternalName as string).LessThanOrEqualTo(endDate)
        ]));

        return new ComparisonResult<TBaseInterface>(this);
    }

    public IsToday(): ComparisonResult<TBaseInterface> {
        const StartToday = new Date(); StartToday.setHours(0, 0, 0, 0);
        const EndToday = new Date(); EndToday.setHours(23, 59, 59, 999);
        return this.IsBetween(StartToday, EndToday);
    }
}






class ComparisonResult<TBaseInterface> extends BaseQuery<TBaseInterface>{
    public Or(): QueryableFields<TBaseInterface> {
        this.query.push(FilterJoinOperator.Or);
        return new QueryableFields<TBaseInterface>(this);
    }

    public And(): QueryableFields<TBaseInterface> {
        this.query.push(FilterJoinOperator.And);
        return new QueryableFields<TBaseInterface>(this);
    }

    public ToString(): string {
        return this.query.join(" ");
    }
}
