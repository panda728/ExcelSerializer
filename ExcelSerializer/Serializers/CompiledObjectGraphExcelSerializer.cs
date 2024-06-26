﻿using System.Buffers;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;

namespace ExcelSerializerLib.Serializers;

internal sealed class CompiledObjectGraphExcelSerializer<T> : IExcelSerializer<T>
{
    private delegate void WriteTitleMethod(ref ExcelFormatter formatter, IBufferWriter<byte> writer, IExcelSerializer?[]? alternateSerializers, T value, ExcelSerializerOptions options);

    private delegate void SerializeMethod(ref ExcelFormatter formatter, IBufferWriter<byte> writer, IExcelSerializer?[]? alternateSerializers, T value, ExcelSerializerOptions options);

    static readonly IExcelSerializer?[]? alternateSerializers;
    static readonly WriteTitleMethod writeTitle;
    static readonly SerializeMethod serialize;
    static readonly bool isReferenceType;

    static CompiledObjectGraphExcelSerializer()
    {
        isReferenceType = !typeof(T).IsValueType;

        var props = typeof(T).GetProperties().Where(p => p.GetIndexParameters().Length == 0);
        var fields = typeof(T).GetFields();
        var members = props.Cast<MemberInfo>().Concat(fields)
            .Where(x => x.GetCustomAttribute<IgnoreExcelSerializeAttribute>() == null)
            .Select((x, i) => new SerializableMemberInfo(x, i))
            .OrderBy(x => x.Order)
            .ThenBy(x => x.Name)
            .ToArray();

        if (members.Any(x => x.ExcelSerializer != null))
        {
            alternateSerializers = members.Select(x => x.ExcelSerializer).ToArray();
        }
        writeTitle = CompileTitleWriter(typeof(T), members);
        serialize = CompileSerializer(typeof(T), members);
    }

    public void WriteTitle(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T value, ExcelSerializerOptions options, string name = "value")
    {
        formatter.EnterAndValidate();
        writeTitle(ref formatter, writer, alternateSerializers, value, options);
        formatter.Exit();
    }

    public void Serialize(ref ExcelFormatter formatter, IBufferWriter<byte> writer, T value, ExcelSerializerOptions options)
    {
        if (isReferenceType)
        {
            if (value == null)
            {
                formatter.WriteEmpty(writer);
                return;
            }
        }

        formatter.EnterAndValidate();
        serialize(ref formatter, writer, alternateSerializers, value, options);
        formatter.Exit();
    }

    static WriteTitleMethod CompileTitleWriter(Type valueType, SerializableMemberInfo[] memberInfos)
    {
        // foreach(members)
        //     options.GetRequiredSerializer<T>() || ((IExcelSerialzier<T>)alternateSerializers[0] .WriteTitle(formatter, value.Foo, options, propertyName)
        var argFormatterRef = Expression.Parameter(typeof(ExcelFormatter).MakeByRefType());
        var argWriter = Expression.Parameter(typeof(IBufferWriter<byte>));
        var argAlternateSerializers = Expression.Parameter(typeof(IExcelSerializer[]));
        var argValue = Expression.Parameter(valueType);
        var argOptions = Expression.Parameter(typeof(ExcelSerializerOptions));
        var foreachBodies = new List<Expression>();

        var i = 0;
        foreach (var memberInfo in memberInfos)
        {
            Expression serializer = memberInfo.ExcelSerializer == null
                ? Expression.Call(argOptions, ReflectionInfos.ExcelSerializerOptions_GetRequiredSerializer(memberInfo.MemberType))
                : Expression.Convert(
                    Expression.ArrayIndex(argAlternateSerializers, Expression.Constant(i, typeof(int))),
                    typeof(IExcelSerializer<>).MakeGenericType(memberInfo.MemberType)
                );

            var callWriteMember = Expression.Call(
                serializer,
                ReflectionInfos.IExcelSerializer_WriteTitle(memberInfo.MemberType),
                argFormatterRef,
                argWriter,
                memberInfo.GetMemberExpression(argValue),
                argOptions,
                Expression.Constant(memberInfo.Name)
            );

            foreachBodies.Add(callWriteMember);
            i++;
        }

        var body = Expression.Block(foreachBodies);
        var lambda = Expression.Lambda<WriteTitleMethod>(body, argFormatterRef, argWriter, argAlternateSerializers, argValue, argOptions);
        return lambda.Compile();
    }

    static SerializeMethod CompileSerializer(Type valueType, SerializableMemberInfo[] memberInfos)
    {
        // foreach(members)
        //   if (value.Foo != null) // reference type || nullable type
        //     options.GetRequiredSerializer<T>() || ((IExcelSerialzier<T>)alternateSerializers[0] .Serialize(formatter, value.Foo, options)
        //   else
        //     formatter.WriteEmpty(writer);
        var argFormatterRef = Expression.Parameter(typeof(ExcelFormatter).MakeByRefType());
        var argWriter = Expression.Parameter(typeof(IBufferWriter<byte>));
        var argAlternateSerializers = Expression.Parameter(typeof(IExcelSerializer[]));
        var argValue = Expression.Parameter(valueType);
        var argOptions = Expression.Parameter(typeof(ExcelSerializerOptions));
        var foreachBodies = new List<Expression>();

        var i = 0;
        foreach (var memberInfo in memberInfos)
        {
            Expression serializer = memberInfo.ExcelSerializer == null
                ? Expression.Call(argOptions, ReflectionInfos.ExcelSerializerOptions_GetRequiredSerializer(memberInfo.MemberType))
                : Expression.Convert(
                    Expression.ArrayIndex(argAlternateSerializers, Expression.Constant(i, typeof(int))),
                    typeof(IExcelSerializer<>).MakeGenericType(memberInfo.MemberType)
                );

            var callSerializer = Expression.Call(
                serializer,
                ReflectionInfos.IExcelSerializer_Serialize(memberInfo.MemberType),
                argFormatterRef,
                argWriter,
                memberInfo.GetMemberExpression(argValue),
                argOptions
            );

            if (!memberInfo.MemberType.IsValueType || memberInfo.MemberType.IsNullable())
            {
                var nullExpr = Expression.Constant(null, memberInfo.MemberType);
                var ifBody = Expression.IfThenElse(
                    Expression.NotEqual(memberInfo.GetMemberExpression(argValue), nullExpr),
                    callSerializer,
                    Expression.Call(argFormatterRef, ReflectionInfos.Writer_Empty, argWriter)
                );
                foreachBodies.Add(ifBody);
            }
            else
            {
                foreachBodies.Add(callSerializer);
            }
            i++;
        }

        var body = Expression.Block(foreachBodies);
        var lambda = Expression.Lambda<SerializeMethod>(body, argFormatterRef, argWriter, argAlternateSerializers, argValue, argOptions);
        return lambda.Compile();

    }

    internal static class ReflectionInfos
    {
        internal static MethodInfo Writer_Empty { get; } = typeof(ExcelFormatter).GetMethod("WriteEmpty")!;
        internal static MethodInfo ExcelSerializerOptions_GetRequiredSerializer(Type type) => typeof(ExcelSerializerOptions).GetMethod("GetRequiredSerializer", 1, Type.EmptyTypes)!.MakeGenericMethod(type);
        internal static MethodInfo IExcelSerializer_Serialize(Type type) => typeof(IExcelSerializer<>).MakeGenericType(type).GetMethod("Serialize")!;
        internal static MethodInfo IExcelSerializer_WriteTitle(Type type) => typeof(IExcelSerializer<>).MakeGenericType(type).GetMethod("WriteTitle")!;
    }
}

internal sealed class SerializableMemberInfo
{
    public string Name { get; }
    public int Order { get; }
    public IExcelSerializer? ExcelSerializer { get; }
    public Type MemberType { get; }
    public MemberInfo MemberInfo { get; }

    public SerializableMemberInfo(MemberInfo member, int i)
    {
        var dataMember = member.GetCustomAttribute<DataMemberAttribute>();

        MemberInfo = member;
        Name = dataMember?.Name ?? member.Name;
        Order = dataMember?.Order ?? i;

        MemberType = member switch
        {
            PropertyInfo pi => pi.PropertyType,
            FieldInfo fi => fi.FieldType,
            _ => throw new InvalidOperationException()
        };

        var serializerAttr = member.GetCustomAttribute<ExcelSerializerAttribute>();
        if (serializerAttr != null)
        {
            serializerAttr.Validate(MemberType);
            ExcelSerializer = (IExcelSerializer?)Activator.CreateInstance(serializerAttr.Type);
        }
    }

    public MemberExpression GetMemberExpression(Expression expression)
    {
        if (MemberInfo is FieldInfo fi)
        {
            return Expression.Field(expression, fi);
        }
        else if (MemberInfo is PropertyInfo pi)
        {
            return Expression.Property(expression, pi);
        }
        throw new InvalidOperationException();
    }
}