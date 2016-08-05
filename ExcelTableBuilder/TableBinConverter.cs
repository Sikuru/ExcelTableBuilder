using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Sikuru.ExcelTableBuilder
{
    /*
     * 번외의 DLL을 생성하지 않기 위한 단순한 클래스 <-> 바이너리 컨버터. Goo*le의 Prot*Buf 보다는 느리지만 아주 느리지도 않습니다.
     * 용량은 적절한 바이트 단위이고, 해당 클래스 구조의 변화에 대한 검증을 하지 않습니다.
     * 프로퍼티 이름에 의한 리플렉션으로 찾아 deserialize를 하기 때문에 클래스 구조가 매칭되지 않는 것에 대한 위험성이 있습니다.
     * 
     * TableBinConvertable 인터페이스를 구현한 클래스도 프로퍼티로 가질 수 있고, depth의 제한은 없습니다.
     * TableBinConvertable 인터페이스를 구현한 클래스와 그 외 지원 가능한 타입은 이하 private static readonly Type 들에 대한 타입만 지원합니다.
     */

    public interface TableBinConvertable
	{
	}

	public static class TableBinConverter
	{
		public const short BinClassDataConverterVersion = 8;
		private static readonly byte[] ZeroLengthBytes = BitConverter.GetBytes((int)0);
		private static readonly Type TypeInt16 = typeof(System.Int16);
		private static readonly Type TypeInt32 = typeof(System.Int32);
		private static readonly Type TypeInt64 = typeof(System.Int64);
		private static readonly Type TypeSingle = typeof(System.Single);
		private static readonly Type TypeDouble = typeof(System.Double);
		private static readonly Type TypeBoolean = typeof(System.Boolean);
		private static readonly Type TypeDateTime = typeof(System.DateTime);
		private static readonly Type TypeString = typeof(System.String);
		private static readonly Type TypeByte = typeof(System.Byte);
		private static readonly Type TypeGuid = typeof(System.Guid);
		private static readonly Type TypeList = typeof(System.Collections.Generic.List<>);
		private const int SizeOfInt16 = sizeof(System.Int16);
		private const int SizeOfInt32 = sizeof(System.Int32);
		private const int SizeOfInt64 = sizeof(System.Int64);
		private const int SizeOfSingle = sizeof(System.Single);
		private const int SizeOfDouble = sizeof(System.Double);
		private const int SizeOfBoolean = sizeof(System.Boolean);
		private const int SizeOfByte = sizeof(System.Byte);
		private const int SizeOfGuid = 16;

		public static T ClassMaker<T>(byte[] byte_data) where T : TableBinConvertable, new()
		{
			var class_object = new T();
			ClassMaker(byte_data, ref class_object);
			return class_object;
		}

		public static T ClassMaker<T>(MemoryStream ms) where T : TableBinConvertable, new()
		{
			var class_object = new T();
			ClassMaker(ms, ref class_object);
			return class_object;
		}

		public static void ClassMaker<T>(byte[] byte_data, ref T class_object) where T : TableBinConvertable
		{
			Type type = typeof(T);
			object raw_object = (object)class_object;
			using (var ms = new MemoryStream(byte_data))
			{
				ClassMaker(ms, type, ref raw_object, false);
			}
		}

		public static void ClassMaker<T>(MemoryStream ms, ref T class_object) where T : TableBinConvertable
		{
			Type type = typeof(T);
			object raw_object = (object)class_object;
			ClassMaker(ms, type, ref raw_object, false);
		}

		public static object ClassMaker(Type type, byte[] byte_data)
		{
			if (type.GetInterface("TableBinConvertable") == null)
			{
				throw new ArgumentException("'TableBinConvertable' CLASS NOT IMPLEMENTED");
			}

			object raw_object = Activator.CreateInstance(type);
			using (var ms = new MemoryStream(byte_data))
			{
				ClassMaker(ms, type, ref raw_object, false);
			}

			return raw_object;
		}

		private static void ClassMaker(MemoryStream ms, Type type, ref object class_object, bool is_inner)
		{
			if (is_inner == false)
			{
				byte[] version_bytes = new byte[SizeOfInt16];
				ms.Read(version_bytes, 0, version_bytes.Length);
				short version = BitConverter.ToInt16(version_bytes, 0);

				if (version != BinClassDataConverterVersion)
				{
					throw new InvalidOperationException("Invalid Data Version");
				}
			}

			var all_pi = type.GetProperties();
			var pi_dic = new SortedDictionary<string, PropertyInfo>(all_pi.ToDictionary(x => x.Name));

			byte[] bytes = new byte[SizeOfGuid];
			foreach (var kv in pi_dic)
			{
				if (kv.Value.PropertyType == TypeByte)
				{
					ms.Read(bytes, 0, SizeOfByte);
					kv.Value.SetValue(class_object, bytes[0], null);
				}
				else if (kv.Value.PropertyType == TypeInt16)
				{
					ms.Read(bytes, 0, SizeOfInt16);
					kv.Value.SetValue(class_object, BitConverter.ToInt16(bytes, 0), null);
				}
				else if (kv.Value.PropertyType == TypeInt32)
				{
					ms.Read(bytes, 0, SizeOfInt32);
					kv.Value.SetValue(class_object, BitConverter.ToInt32(bytes, 0), null);
				}
				else if (kv.Value.PropertyType == TypeInt64)
				{
					ms.Read(bytes, 0, SizeOfInt64);
					kv.Value.SetValue(class_object, BitConverter.ToInt64(bytes, 0), null);
				}
				else if (kv.Value.PropertyType == TypeSingle)
				{
					ms.Read(bytes, 0, SizeOfSingle);
					kv.Value.SetValue(class_object, BitConverter.ToSingle(bytes, 0), null);
				}
				else if (kv.Value.PropertyType == TypeDouble)
				{
					ms.Read(bytes, 0, SizeOfDouble);
					kv.Value.SetValue(class_object, BitConverter.ToDouble(bytes, 0), null);
				}
				else if (kv.Value.PropertyType == TypeBoolean)
				{
					ms.Read(bytes, 0, SizeOfBoolean);
					kv.Value.SetValue(class_object, BitConverter.ToBoolean(bytes, 0), null);
				}
				else if (kv.Value.PropertyType == TypeDateTime)
				{
					ms.Read(bytes, 0, SizeOfInt64);

					long time_long = BitConverter.ToInt64(bytes, 0);
					kv.Value.SetValue(class_object, DateTime.FromBinary(time_long), null);
				}
				else if (kv.Value.PropertyType == TypeGuid)
				{
					ms.Read(bytes, 0, SizeOfGuid);
					kv.Value.SetValue(class_object, new Guid(bytes), null);
				}
				else if (kv.Value.PropertyType.IsEnum)
				{
					var enum_type = Enum.GetUnderlyingType(kv.Value.PropertyType);
					if (enum_type == TypeByte)
					{
						ms.Read(bytes, 0, SizeOfByte);
						kv.Value.SetValue(class_object, Enum.ToObject(kv.Value.PropertyType, bytes[0]), null);
					}
					else if (enum_type == TypeInt16)
					{
						ms.Read(bytes, 0, SizeOfInt16);
						kv.Value.SetValue(class_object, Enum.ToObject(kv.Value.PropertyType, BitConverter.ToInt16(bytes, 0)), null);
					}
					else if (enum_type == TypeInt32)
					{
						ms.Read(bytes, 0, SizeOfInt32);
						kv.Value.SetValue(class_object, Enum.ToObject(kv.Value.PropertyType, BitConverter.ToInt32(bytes, 0)), null);
					}
					else if (enum_type == TypeInt64)
					{
						ms.Read(bytes, 0, SizeOfInt64);
						kv.Value.SetValue(class_object, Enum.ToObject(kv.Value.PropertyType, BitConverter.ToInt64(bytes, 0)), null);
					}
					else
					{
						throw new NotSupportedException("Not supported enum underlying type");
					}
				}
				else if (kv.Value.PropertyType == TypeString)
				{
					ms.Read(bytes, 0, SizeOfInt32);
					int string_length = BitConverter.ToInt32(bytes, 0);
					if (string_length > 0)
					{
						byte[] string_bytes = new byte[string_length];
						ms.Read(string_bytes, 0, string_length);

						kv.Value.SetValue(class_object, Encoding.UTF8.GetString(string_bytes), null);
					}
					else
					{
						kv.Value.SetValue(class_object, String.Empty, null);
					}
				}
				else
				{
					ms.Read(bytes, 0, SizeOfInt32);
					int length = BitConverter.ToInt32(bytes, 0);

					if (length > 0)
					{
						if (kv.Value.PropertyType.IsArray)
						{
							Type array_type = kv.Value.PropertyType.GetElementType();

							if (array_type == TypeByte)
							{
								byte[] binary_bytes = new byte[length];
								ms.Read(binary_bytes, 0, length);

								kv.Value.SetValue(class_object, binary_bytes, null);
							}
							else
							{
								Array array = Array.CreateInstance(array_type, length);
								for (int i = 0; i < length; ++i)
								{
									object child_class;

									if (array_type == TypeByte)
									{
										ms.Read(bytes, 0, SizeOfByte);
										child_class = bytes[0];
									}
									else if (array_type == TypeInt16)
									{
										ms.Read(bytes, 0, SizeOfInt16);
										child_class = BitConverter.ToInt16(bytes, 0);
									}
									else if (array_type == TypeInt32)
									{
										ms.Read(bytes, 0, SizeOfInt32);
										child_class = BitConverter.ToInt32(bytes, 0);
									}
									else if (array_type == TypeInt64)
									{
										ms.Read(bytes, 0, SizeOfInt64);
										child_class = BitConverter.ToInt64(bytes, 0);
									}
									else if (array_type == TypeSingle)
									{
										ms.Read(bytes, 0, SizeOfSingle);
										child_class = BitConverter.ToSingle(bytes, 0);
									}
									else if (array_type == TypeDouble)
									{
										ms.Read(bytes, 0, SizeOfDouble);
										child_class = BitConverter.ToDouble(bytes, 0);
									}
									else if (array_type == TypeBoolean)
									{
										ms.Read(bytes, 0, SizeOfBoolean);
										child_class = BitConverter.ToBoolean(bytes, 0);
									}
									else if (array_type == TypeDateTime)
									{
										ms.Read(bytes, 0, SizeOfInt64);

										long time_long = BitConverter.ToInt64(bytes, 0);
										child_class = DateTime.FromBinary(time_long);
									}
									else if (array_type == TypeGuid)
									{
										ms.Read(bytes, 0, SizeOfGuid);
										child_class = new Guid(bytes);
									}
									else if (array_type.IsEnum)
									{
										var enum_type = Enum.GetUnderlyingType(array_type);
										if (enum_type == TypeByte)
										{
											ms.Read(bytes, 0, SizeOfByte);
											child_class = Enum.ToObject(array_type, bytes[0]);
										}
										else if (enum_type == TypeInt16)
										{
											ms.Read(bytes, 0, SizeOfInt16);
											child_class = Enum.ToObject(array_type, BitConverter.ToInt16(bytes, 0));
										}
										else if (enum_type == TypeInt32)
										{
											ms.Read(bytes, 0, SizeOfInt32);
											child_class = Enum.ToObject(array_type, BitConverter.ToInt32(bytes, 0));
										}
										else if (enum_type == TypeInt64)
										{
											ms.Read(bytes, 0, SizeOfInt64);
											child_class = Enum.ToObject(array_type, BitConverter.ToInt64(bytes, 0));
										}
										else
										{
											throw new NotSupportedException("Not supported enum underlying type");
										}
									}
									else if (array_type == TypeString)
									{
										ms.Read(bytes, 0, SizeOfInt32);
										int string_length = BitConverter.ToInt32(bytes, 0);
										if (string_length > 0)
										{
											byte[] string_bytes = new byte[string_length];
											ms.Read(string_bytes, 0, string_length);

											child_class = Encoding.UTF8.GetString(string_bytes);
										}
										else
										{
											child_class = String.Empty;
										}
									}
									else
									{
										child_class = Activator.CreateInstance(array_type);
										ClassMaker(ms, child_class.GetType(), ref child_class, true);
									}

									array.SetValue(child_class, i);
								}

								kv.Value.SetValue(class_object, array, null);
							}
						}
						else if (kv.Value.PropertyType.IsGenericType && kv.Value.PropertyType.GetGenericTypeDefinition() == TypeList)
						{
							Type item_type = kv.Value.PropertyType.GetGenericArguments()[0];
							Type[] list_args = { item_type };

							object list = Activator.CreateInstance(TypeList.MakeGenericType(list_args), null);

							for (int i = 0; i < length; ++i)
							{
								object child_class;

								if (item_type == TypeByte)
								{
									throw new NotSupportedException("List<byte> is not supported");
								}
								else if (item_type == TypeInt16)
								{
									ms.Read(bytes, 0, SizeOfInt16);
									child_class = BitConverter.ToInt16(bytes, 0);
								}
								else if (item_type == TypeInt32)
								{
									ms.Read(bytes, 0, SizeOfInt32);
									child_class = BitConverter.ToInt32(bytes, 0);
								}
								else if (item_type == TypeInt64)
								{
									ms.Read(bytes, 0, SizeOfInt64);
									child_class = BitConverter.ToInt64(bytes, 0);
								}
								else if (item_type == TypeSingle)
								{
									ms.Read(bytes, 0, SizeOfSingle);
									child_class = BitConverter.ToSingle(bytes, 0);
								}
								else if (item_type == TypeDouble)
								{
									ms.Read(bytes, 0, SizeOfDouble);
									child_class = BitConverter.ToDouble(bytes, 0);
								}
								else if (item_type == TypeBoolean)
								{
									ms.Read(bytes, 0, SizeOfBoolean);
									child_class = BitConverter.ToBoolean(bytes, 0);
								}
								else if (item_type == TypeDateTime)
								{
									ms.Read(bytes, 0, SizeOfInt64);

									long time_long = BitConverter.ToInt64(bytes, 0);
									child_class = DateTime.FromBinary(time_long);
								}
								else if (item_type == TypeGuid)
								{
									ms.Read(bytes, 0, SizeOfGuid);
									child_class = new Guid(bytes);
								}
								else if (item_type.IsEnum)
								{
									var enum_type = Enum.GetUnderlyingType(item_type);
									if (enum_type == TypeByte)
									{
										ms.Read(bytes, 0, SizeOfByte);
										child_class = Enum.ToObject(item_type, bytes[0]);
									}
									else if (enum_type == TypeInt16)
									{
										ms.Read(bytes, 0, SizeOfInt16);
										child_class = Enum.ToObject(item_type, BitConverter.ToInt16(bytes, 0));
									}
									else if (enum_type == TypeInt32)
									{
										ms.Read(bytes, 0, SizeOfInt32);
										child_class = Enum.ToObject(item_type, BitConverter.ToInt32(bytes, 0));
									}
									else if (enum_type == TypeInt64)
									{
										ms.Read(bytes, 0, SizeOfInt64);
										child_class = Enum.ToObject(item_type, BitConverter.ToInt64(bytes, 0));
									}
									else
									{
										throw new NotSupportedException("Not supported enum underlying type");
									}
								}
								else if (item_type == TypeString)
								{
									ms.Read(bytes, 0, SizeOfInt32);
									int string_length = BitConverter.ToInt32(bytes, 0);
									if (string_length > 0)
									{
										byte[] string_bytes = new byte[string_length];
										ms.Read(string_bytes, 0, string_length);

										child_class = Encoding.UTF8.GetString(string_bytes);
									}
									else
									{
										child_class = String.Empty;
									}
								}
								else
								{
									child_class = Activator.CreateInstance(item_type);
									ClassMaker(ms, child_class.GetType(), ref child_class, true);
								}

								((IList)list).Add(child_class);
							}

							kv.Value.SetValue(class_object, list, null);
						}
						else
						{
							object child_class = Activator.CreateInstance(kv.Value.PropertyType);
							ClassMaker(ms, child_class.GetType(), ref child_class, true);
							kv.Value.SetValue(class_object, child_class, null);
						}
					}
					else
					{
						if (kv.Value.PropertyType.IsGenericType && kv.Value.PropertyType.GetGenericTypeDefinition() == TypeList)
						{
							Type item_type = kv.Value.PropertyType.GetGenericArguments()[0];
							Type[] list_args = { item_type };

							object list = Activator.CreateInstance(TypeList.MakeGenericType(list_args), null);
							kv.Value.SetValue(class_object, list, null);
						}
						else
						{
							kv.Value.SetValue(class_object, null, null);
						}
					}
				}
			}
		}

		public static byte[] BinMaker<T>(T class_object) where T : TableBinConvertable
		{
			Type type = typeof(T);
			return BinMaker(type, class_object, false);
		}

		public static byte[] BinMaker(Type type, object class_object)
		{
			return BinMaker(type, class_object, false);
		}

		private static byte[] BinMaker(Type type, object class_object, bool is_inner)
		{
			if (type.GetInterface("TableBinConvertable") == null)
			{
				throw new ArgumentException("'TableBinConvertable' CLASS NOT IMPLEMENTED");
			}

			var all_pi = type.GetProperties();
			var pi_dic = new SortedDictionary<string, PropertyInfo>(all_pi.ToDictionary(x => x.Name));

			using (var ms = new MemoryStream())
			{
				if (is_inner == false)
				{
					byte[] version_bytes = BitConverter.GetBytes(BinClassDataConverterVersion);
					ms.Write(version_bytes, 0, version_bytes.Length);
				}

				foreach (var kv in pi_dic)
				{
					if (kv.Value.PropertyType.IsArray)
					{
						Array array = (Array)kv.Value.GetValue(class_object, null);
						if (array != null)
						{
							byte[] array_count_bytes = BitConverter.GetBytes((int)array.Length);
							ms.Write(array_count_bytes, 0, array_count_bytes.Length);

							var array_type = kv.Value.PropertyType.GetElementType();

							if (array_type == TypeByte)
							{
								byte[] bytes = (byte[])array;
								ms.Write(bytes, 0, bytes.Length);
							}
							else
							{
								foreach (var array_object in array)
								{
									if (array_type == TypeInt16)
									{
										ms.Write(BitConverter.GetBytes((short)array_object), 0, SizeOfInt16);
									}
									else if (array_type == TypeInt32)
									{
										ms.Write(BitConverter.GetBytes((int)array_object), 0, SizeOfInt32);
									}
									else if (array_type == TypeInt64)
									{
										ms.Write(BitConverter.GetBytes((long)array_object), 0, SizeOfInt64);
									}
									else if (array_type == TypeSingle)
									{
										ms.Write(BitConverter.GetBytes((float)array_object), 0, SizeOfSingle);
									}
									else if (array_type == TypeDouble)
									{
										ms.Write(BitConverter.GetBytes((double)array_object), 0, SizeOfDouble);
									}
									else if (array_type == TypeBoolean)
									{
										ms.Write(BitConverter.GetBytes((bool)array_object), 0, SizeOfBoolean);
									}
									else if (array_type == TypeDateTime)
									{
										ms.Write(BitConverter.GetBytes(((DateTime)array_object).ToBinary()), 0, SizeOfInt64);
									}
									else if (array_type == TypeGuid)
									{
										ms.Write(((Guid)array_object).ToByteArray(), 0, SizeOfGuid);
									}
									else if (array_type.IsEnum)
									{
										var enum_type = Enum.GetUnderlyingType(array_type);
										if (enum_type == TypeByte)
										{
											ms.Write(BitConverter.GetBytes((byte)array_object), 0, SizeOfByte);
										}
										else if (enum_type == TypeInt16)
										{
											ms.Write(BitConverter.GetBytes((short)array_object), 0, SizeOfInt16);
										}
										else if (enum_type == TypeInt32)
										{
											ms.Write(BitConverter.GetBytes((int)array_object), 0, SizeOfInt32);
										}
										else if (enum_type == TypeInt64)
										{
											ms.Write(BitConverter.GetBytes((long)array_object), 0, SizeOfInt64);
										}
										else
										{
											throw new NotSupportedException("Not supported enum underlying type");
										}
									}
									else if (array_type == TypeString)
									{
										string string_data = array_object == null ? "" : array_object.ToString();
										if (!String.IsNullOrEmpty(string_data))
										{
											byte[] string_bytes = Encoding.UTF8.GetBytes(string_data);
											if (string_bytes.Length > int.MaxValue)
											{
												throw new ArgumentOutOfRangeException(kv.Value.Name);
											}
											byte[] length_bytes = BitConverter.GetBytes((int)string_bytes.Length);

											ms.Write(length_bytes, 0, length_bytes.Length);
											ms.Write(string_bytes, 0, string_bytes.Length);
										}
										else
										{
											ms.Write(ZeroLengthBytes, 0, ZeroLengthBytes.Length);
										}
									}
									else
									{
										byte[] bytes = BinMaker(array_type, array_object, true);
										ms.Write(bytes, 0, bytes.Length);
									}
								}
							}
						}
						else
						{
							ms.Write(ZeroLengthBytes, 0, ZeroLengthBytes.Length);
						}
					}
					else if (kv.Value.PropertyType.IsGenericType && kv.Value.PropertyType.GetGenericTypeDefinition() == TypeList)
					{
						var list = kv.Value.GetValue(class_object, null);

						ICollection collection = (ICollection)list;
						if (collection != null)
						{
							Type item_type = kv.Value.PropertyType.GetGenericArguments()[0];
							int count = collection.Count;

							byte[] array_count_bytes = BitConverter.GetBytes((int)count);
							ms.Write(array_count_bytes, 0, array_count_bytes.Length);

							foreach (object list_item in (IEnumerable)list)
							{
								if (item_type == TypeByte)
								{
									throw new NotSupportedException("List<byte> not supported");
								}
								else if (item_type == TypeInt16)
								{
									ms.Write(BitConverter.GetBytes((short)list_item), 0, SizeOfInt16);
								}
								else if (item_type == TypeInt32)
								{
									ms.Write(BitConverter.GetBytes((int)list_item), 0, SizeOfInt32);
								}
								else if (item_type == TypeInt64)
								{
									ms.Write(BitConverter.GetBytes((long)list_item), 0, SizeOfInt64);
								}
								else if (item_type == TypeSingle)
								{
									ms.Write(BitConverter.GetBytes((float)list_item), 0, SizeOfSingle);
								}
								else if (item_type == TypeDouble)
								{
									ms.Write(BitConverter.GetBytes((double)list_item), 0, SizeOfDouble);
								}
								else if (item_type == TypeBoolean)
								{
									ms.Write(BitConverter.GetBytes((bool)list_item), 0, SizeOfBoolean);
								}
								else if (item_type == TypeDateTime)
								{
									ms.Write(BitConverter.GetBytes(((DateTime)list_item).ToBinary()), 0, SizeOfInt64);
								}
								else if (item_type == TypeGuid)
								{
									ms.Write(((Guid)list_item).ToByteArray(), 0, SizeOfGuid);
								}
								else if (item_type.IsEnum)
								{
									var enum_type = Enum.GetUnderlyingType(item_type);
									if (enum_type == TypeByte)
									{
										ms.Write(BitConverter.GetBytes((byte)list_item), 0, SizeOfByte);
									}
									else if (enum_type == TypeInt16)
									{
										ms.Write(BitConverter.GetBytes((short)list_item), 0, SizeOfInt16);
									}
									else if (enum_type == TypeInt32)
									{
										ms.Write(BitConverter.GetBytes((int)list_item), 0, SizeOfInt32);
									}
									else if (enum_type == TypeInt64)
									{
										ms.Write(BitConverter.GetBytes((long)list_item), 0, SizeOfInt64);
									}
									else
									{
										throw new NotSupportedException("Not supported enum underlying type");
									}
								}
								else if (item_type == TypeString)
								{
									string string_data = list_item == null ? "" : list_item.ToString();
									if (!String.IsNullOrEmpty(string_data))
									{
										byte[] string_bytes = Encoding.UTF8.GetBytes(string_data);
										if (string_bytes.Length > int.MaxValue)
										{
											throw new ArgumentOutOfRangeException(kv.Value.Name);
										}
										byte[] length_bytes = BitConverter.GetBytes((int)string_bytes.Length);

										ms.Write(length_bytes, 0, length_bytes.Length);
										ms.Write(string_bytes, 0, string_bytes.Length);
									}
									else
									{
										ms.Write(ZeroLengthBytes, 0, ZeroLengthBytes.Length);
									}
								}
								else
								{
									byte[] bytes = BinMaker(item_type, list_item, true);
									ms.Write(bytes, 0, bytes.Length);
								}
							}
						}
						else
						{
							ms.Write(ZeroLengthBytes, 0, ZeroLengthBytes.Length);
						}
					}
					else if (kv.Value.PropertyType == TypeByte)
					{
						byte[] bytes = new[] { (byte)kv.Value.GetValue(class_object, null) };
						ms.Write(bytes, 0, SizeOfByte);
					}
					else if (kv.Value.PropertyType == TypeInt16)
					{
						ms.Write(BitConverter.GetBytes((short)kv.Value.GetValue(class_object, null)), 0, SizeOfInt16);
					}
					else if (kv.Value.PropertyType == TypeInt32)
					{
						ms.Write(BitConverter.GetBytes((int)kv.Value.GetValue(class_object, null)), 0, SizeOfInt32);
					}
					else if (kv.Value.PropertyType == TypeInt64)
					{
						ms.Write(BitConverter.GetBytes((long)kv.Value.GetValue(class_object, null)), 0, SizeOfInt64);
					}
					else if (kv.Value.PropertyType == TypeSingle)
					{
						ms.Write(BitConverter.GetBytes((float)kv.Value.GetValue(class_object, null)), 0, SizeOfSingle);
					}
					else if (kv.Value.PropertyType == TypeDouble)
					{
						ms.Write(BitConverter.GetBytes((double)kv.Value.GetValue(class_object, null)), 0, SizeOfDouble);
					}
					else if (kv.Value.PropertyType == TypeBoolean)
					{
						ms.Write(BitConverter.GetBytes((bool)kv.Value.GetValue(class_object, null)), 0, SizeOfBoolean);
					}
					else if (kv.Value.PropertyType == TypeDateTime)
					{
						ms.Write(BitConverter.GetBytes(((DateTime)kv.Value.GetValue(class_object, null)).ToBinary()), 0, SizeOfInt64);
					}
					else if (kv.Value.PropertyType == TypeGuid)
					{
						ms.Write(((Guid)kv.Value.GetValue(class_object, null)).ToByteArray(), 0, SizeOfGuid);
					}
					else if (kv.Value.PropertyType.IsEnum)
					{
						var enum_type = Enum.GetUnderlyingType(kv.Value.PropertyType);
						if (enum_type == TypeByte)
						{
							ms.Write(BitConverter.GetBytes((byte)kv.Value.GetValue(class_object, null)), 0, SizeOfByte);
						}
						else if (enum_type == TypeInt16)
						{
							ms.Write(BitConverter.GetBytes((short)kv.Value.GetValue(class_object, null)), 0, SizeOfInt16);
						}
						else if (enum_type == TypeInt32)
						{
							ms.Write(BitConverter.GetBytes((int)kv.Value.GetValue(class_object, null)), 0, SizeOfInt32);
						}
						else if (enum_type == TypeInt64)
						{
							ms.Write(BitConverter.GetBytes((long)kv.Value.GetValue(class_object, null)), 0, SizeOfInt64);
						}
						else
						{
							throw new NotSupportedException("Not supported enum underlying type");
						}
					}
					else if (kv.Value.PropertyType == TypeString)
					{
						object string_object = kv.Value.GetValue(class_object, null);
						string string_data = string_object == null ? "" : string_object.ToString();
						if (!String.IsNullOrEmpty(string_data))
						{
							byte[] string_bytes = Encoding.UTF8.GetBytes(string_data);
							if (string_bytes.Length > int.MaxValue)
							{
								throw new ArgumentOutOfRangeException(kv.Value.Name);
							}
							byte[] length_bytes = BitConverter.GetBytes((int)string_bytes.Length);

							ms.Write(length_bytes, 0, length_bytes.Length);
							ms.Write(string_bytes, 0, string_bytes.Length);
						}
						else
						{
							ms.Write(ZeroLengthBytes, 0, ZeroLengthBytes.Length);
						}
					}
					else
					{
						object child_class = kv.Value.GetValue(class_object, null);
						if (child_class != null)
						{
							byte[] bytes = BinMaker(child_class.GetType(), child_class, true);
							if (bytes.Length > int.MaxValue)
							{
								throw new ArgumentOutOfRangeException(kv.Value.Name);
							}
							byte[] length_bytes = BitConverter.GetBytes((int)bytes.Length);

							ms.Write(length_bytes, 0, length_bytes.Length);
							ms.Write(bytes, 0, bytes.Length);
						}
						else
						{
							ms.Write(ZeroLengthBytes, 0, ZeroLengthBytes.Length);
						}
					}
				}

				return ms.ToArray();
			}
		}
	}
}
