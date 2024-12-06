using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab5
{
    public class Program
    {
        static void Main(string[] args)
        {
            const string logFileName = "log.txt";

            bool createNewFile = false;
            while (true)
            {
                Console.Write("Писать логи в новый файл или в старый? (new/old): ");
                string input = Console.ReadLine()?.Trim().ToLower();

                if (input == "new")
                {
                    createNewFile = true;
                    break;
                }
                else if (input == "old")
                {
                    createNewFile = false;
                    break;
                }
                else
                {
                    Console.WriteLine("Ошибка: введите 'new' для нового или 'old' для старого файла.");
                }
            }

            logger.InitializeLogger(logFileName, !createNewFile);

            // Чтение базы данных из Excel-файла
            var operations = new Operations("LR5-var16.xls");
            logger.LogAction("Начата работа программы.");

            // Вывод меню
            while (true)
            {
                Console.WriteLine("Выберите действие:");
                Console.WriteLine("1. Просмотр базы данных");
                Console.WriteLine("2. Удаление элементов");
                Console.WriteLine("3. Корректировка элементов");
                Console.WriteLine("4. Добавление элементов");
                Console.WriteLine("5. Выполнение запросов");
                Console.WriteLine("0. Выход");

                int choice = int.Parse(Console.ReadLine());

                switch (choice)
                {
                    case 1:
                        Console.WriteLine("Выберите таблицу базы данных(1-4)");
                        Console.WriteLine("1. Альбомы");
                        Console.WriteLine("2. Артисты");
                        Console.WriteLine("3. Треки");
                        Console.WriteLine("4. Жанры");
                        logger.LogAction("Выбран пункт 'Просмотр базы данных'.");

                        int viewChoice = operations.GetIntInput("Введите номер таблицы: ", 1, 4);

                        switch (viewChoice)
                        {
                            case 1:
                                Console.WriteLine("Альбомы:");
                                operations.DisplayAlbums();
                                logger.LogAction("Пользователь просмотрел таблицу 'Альбомы'.");
                                break;
                            case 2:
                                Console.WriteLine("Артисты:");
                                operations.DisplayArtists();
                                logger.LogAction("Пользователь просмотрел таблицу 'Артисты'.");
                                break;
                            case 3:
                                Console.WriteLine("Треки:");
                                operations.DisplayTracks();
                                logger.LogAction("Пользователь просмотрел таблицу 'Треки'.");
                                break;
                            case 4:
                                Console.WriteLine("Жанры:");
                                operations.DisplayGenres();
                                logger.LogAction("Пользователь просмотрел таблицу 'Жанры'.");
                                break;
                            default:
                                Console.WriteLine("Неверный выбор. Попробуйте еще раз.");
                                logger.LogAction("Ошибка ввода при выборе категории просмотра.");
                                break;

                        }
                        break;
                    case 2:
                        Console.WriteLine("Из какой таблицы вы хотите удалить элемент?");
                        Console.WriteLine("1. Альбомы");
                        Console.WriteLine("2. Артисты");
                        Console.WriteLine("3. Треки");
                        Console.WriteLine("4. Жанры");
                        logger.LogAction("Выбран пункт 'Удаление элементов'.");

                        int delChoice = operations.GetIntInput("Введите номер таблицы: ", 1, 4);

                        int id = operations.GetIntInput("Введите ID элемента для удаления: ", 1, int.MaxValue);

                        switch (delChoice)
                        {
                            case 1:
                                operations.DeleteAlbumById(id);
                                break;
                            case 2:
                                operations.DeleteArtistById(id);
                                break;
                            case 3:
                                operations.DeleteTrackById(id);
                                break;
                            case 4:
                                operations.DeleteGenreById(id);
                                break;
                            default:
                                Console.WriteLine("Неверный выбор. Попробуйте еще раз.");
                                logger.LogAction("Ошибка ввода при выборе категории удаления.");
                                break;
                        }
                        break;
                    case 3:
                        Console.WriteLine("В какой таблице вы хотите корректировать элементы?");
                        Console.WriteLine("1. Альбомы");
                        Console.WriteLine("2. Артисты");
                        Console.WriteLine("3. Треки");
                        Console.WriteLine("4. Жанры");
                        logger.LogAction("Выбран пункт 'Корректировка элементов'.");
                        int updchoice = operations.GetIntInput("Выберите таблицу для изменения (1-4):", 1, 4);

                        switch (updchoice)
                        {
                            case 1:
                                operations.UpdateAlbum();
                                break;
                            case 2:
                                operations.UpdateArtist();
                                break;
                            case 3:
                                operations.UpdateTrack();
                                break;
                            case 4:
                                operations.UpdateGenre();
                                break;
                            default:
                                Console.WriteLine("Неверный выбор. Попробуйте еще раз.");
                                logger.LogAction("Ошибка ввода при выборе категории корректирования.");
                                break;
                        }
                        break;
                    case 4:
                        Console.WriteLine("В какую таблицу вы хотите добавить элемент?");
                        Console.WriteLine("1. Альбомы");
                        Console.WriteLine("2. Артисты");
                        Console.WriteLine("3. Треки");
                        Console.WriteLine("4. Жанры");
                        logger.LogAction("Выбран пункт 'Добавление элементов'.");

                        int addchoice = operations.GetIntInput("Выберите таблицу для добавления (1-4):", 1, 4);

                        switch (addchoice)
                        {
                            case 1:
                                operations.AddAlbum();
                                break;
                            case 2:
                                operations.AddArtist();
                                break;
                            case 3:
                                operations.AddTrack();
                                break;
                            case 4:
                                operations.AddGenre();
                                break;
                            default:
                                Console.WriteLine("Неверный выбор. Попробуйте еще раз.");
                                logger.LogAction("Ошибка ввода при выборе категории добавления.");
                                break;
                        }
                        break;
                    case 5:
                        Console.WriteLine("Доступные запросы");
                        Console.WriteLine("1. Вывести количество треков которые стоят 150 рублей и дороже");
                        Console.WriteLine("2. Вывести самые 'дорогие' треки из каждого жанра");
                        Console.WriteLine("3. Вывести все альбомы и треки внутри них исполнителя Оззи Озборн");
                        Console.WriteLine("4. Вывести исполнителя в жанре Metal с наименьшим суммарным размером песен в этом жанре");
                        logger.LogAction("Выбран пункт 'Выполнение запросов'.");

                        int reqchoice = operations.GetIntInput("Введите номер запроса(1-4):", 1, 4);

                        switch (reqchoice)
                        {
                            case 1:
                                operations.GetTrackCountWithPriceGreaterThan150();
                                logger.LogAction("Выбран запрос 'Вывести количество треков которые стоят 150 руйблей и дороже'.");
                                break;
                            case 2:
                                operations.DisplayTracksWithHighestPriceByGenre();
                                logger.LogAction("Выбран запрос 'Вывести самые 'дорогие' треки из каждого жанра'.");
                                break;
                            case 3:
                                operations.DisplayAlbumsAndTracksByOzzyOsbourne();
                                logger.LogAction("Выбран запрос 'Вывести все альбомы и треки внутри них исполнителя Оззи Озборн'.");
                                break;
                            case 4:
                                operations.FindMetalArtistWithSmallestTotalTrackSize();
                                logger.LogAction("Выбран запрос 'Вывести исполнителя в жанре Metal с наименьшим суммарным размером песен в этом жанре'.");
                                break;
                            default:
                                Console.WriteLine("Неверный выбор. Попробуйте еще раз.");
                                logger.LogAction("Ошибка ввода при выборе категории запросов.");
                                break;
                        }
                        break;
                    case 0:
                        logger.LogAction("Завершена работа программы.");
                        return;
                    default:
                        Console.WriteLine("Неверный выбор. Попробуйте еще раз.");
                        logger.LogAction("Ошибка ввода при выборе меню.");
                        break;
                }
            }
        }
    }
}
